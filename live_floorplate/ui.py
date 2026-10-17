"""Additive Streamlit page; no changes to existing Report Tools."""
import hashlib
import json
from datetime import datetime, timedelta, timezone
import fitz
import streamlit as st
import streamlit.components.v1 as components
from .api import sync
from .layout import parse_layout, level_title, matched
from .model import FloorplateError
from .render import Style, reference_style, html_view, powerpoint
from .sample import demonstration

KEY = 'live_floorplate_snapshot'


def synchronize_into(state, config, layout):
    state.pop(KEY,None)
    if config.get('project') != 'Hurstville Rise':
        raise FloorplateError('Wrong configured project')
    data = sync(config)
    matched(data,layout)
    state[KEY] = data


def combine_pdfs(files):
    with fitz.open() as combined:
        for file in files:
            with fitz.open(stream=file.getvalue(),filetype='pdf') as source:
                if source.needs_pass or len(source)+len(combined) > 100:
                    raise FloorplateError('Unsupported PDF')
                combined.insert_pdf(source)
        return combined.tobytes()


@st.fragment(run_every=5)
def show_snapshot(layout, pdf, style):
    data = st.session_state.get(KEY)
    if data is None:
        st.info('No current validated snapshot. Synchronize before viewing prices.')
        return
    st.caption(f'Last synced: {data.synced_at.isoformat()} (UTC). As-of snapshot, not a streaming feed.')
    if not data.fresh():
        st.warning('Snapshot expired. Prices and PowerPoint export are hidden. Synchronize again.')
        return
    try:
        level = st.selectbox('Floor',range(9),format_func=level_title,key='live_floorplate_level')
        pairs = matched(data,layout,level)
        html = html_view(data,layout,pdf,level,style)
        # Hide the iframe at expiry even while the surrounding app is idle.
        remaining = int((data.synced_at+timedelta(seconds=300)-datetime.now(timezone.utc)).total_seconds()*1000)
        html = html.replace('</body>',f'<script>setTimeout(()=>{{document.body.textContent="Snapshot expired. Synchronize again."}},Math.max(0,{remaining}));</script></body>')
        if hasattr(st,'iframe'):
            st.iframe(html,height=720)
        else:
            components.html(html,height=720,scrolling=True)
        st.caption(f'{len(pairs)} Available apartments on this floor. Areas shown as Internal + External.')
        st.download_button('Download as-of PowerPoint — Ground through Level 8',
                           powerpoint(data,layout,pdf,style),file_name='Hurstville_Rise_Floorplates.pptx',
                           mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                           key='live_floorplate_download')
    except Exception:
        st.error('Unable to render a verified floorplate. Check every Available Lot mapping, PDF and label style. Prices are hidden.')


def render_page():
    st.header('Live Floorplate — Hurstville Rise')
    mode = st.radio('Data source',['Synthetic demonstration','Configured Agent Portal API'],key='live_floorplate_mode')
    if st.session_state.get('live_floorplate_previous_mode') != mode:
        st.session_state.pop(KEY,None)
        st.session_state['live_floorplate_previous_mode'] = mode
    if mode == 'Synthetic demonstration':
        st.warning('SAMPLE DATA AND DRAWINGS — not an Arcanite response, real inventory, confirmed Hurstville layout or approved Landmark artwork.')
        data,layout,pdf = demonstration()
        refresh = st.button('Refresh sample snapshot')
        if KEY not in st.session_state or refresh:
            st.session_state[KEY] = data
        style = Style()
    else:
        st.info('Use only the company-approved environment. Credentials belong in server-side Streamlit Secrets; never upload them here.')
        originals = st.file_uploader('Original architectural PDF floorplates (ordered; combined page indices start at 0)',
                                     type=['pdf'],accept_multiple_files=True,key='live_floorplate_pdfs')
        mapping = st.file_uploader('Owner-reviewed lot-to-position mapping JSON',type=['json'],key='live_floorplate_mapping')
        reference = st.file_uploader('Reference Landmark PowerPoint for label styling',type=['pptx'],key='live_floorplate_style')
        clean = st.checkbox('PDF backgrounds contain no old price/availability labels, and all Lot positions are owner-reviewed.',key='live_floorplate_clean')
        context = hashlib.sha256(b''.join(f.getvalue() for f in originals)+(mapping.getvalue() if mapping else b'')+
                                 (reference.getvalue() if reference else b'')+str(clean).encode()).hexdigest()
        if st.session_state.get('live_floorplate_context') != context:
            st.session_state.pop(KEY,None)
            st.session_state['live_floorplate_context'] = context
        if not (originals and mapping and reference and clean):
            st.info('Provide the original PDFs, confirmed mapping and reference styling before synchronizing. No inferred positions are used.')
            return
        try:
            pdf = combine_pdfs(originals)
            with fitz.open(stream=pdf,filetype='pdf') as doc:
                layout = parse_layout(json.loads(mapping.getvalue()),len(doc))
            style = reference_style(reference.getvalue())
        except Exception:
            st.session_state.pop(KEY,None)
            st.error('Floorplate assets or mapping failed validation. No prices will be displayed.')
            return
        if st.button('Synchronize authorized Agent Portal inventory',type='primary'):
            # Invalidate BEFORE networking: failed refresh cannot retain old prices.
            st.session_state.pop(KEY,None)
            try:
                config = dict(st.secrets['arcanite'])
                synchronize_into(st.session_state,config,layout)
            except Exception:
                # Do not expose URLs, headers, upstream errors or response bodies.
                st.error('Synchronization failed or completeness/access could not be verified. Prices and export are unavailable. Ask the authorized operator to verify the contract.')
    show_snapshot(layout,pdf,style)
