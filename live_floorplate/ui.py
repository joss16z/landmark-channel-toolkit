"""Additive Streamlit page; no changes to existing Report Tools."""
import hashlib
import json
import os
from pathlib import Path
from datetime import datetime, timedelta, timezone
import fitz
import streamlit as st
import streamlit.components.v1 as components
from .api import sync
from .connection_test import connection_report, compare_project_list_requests
from .layout import parse_layout, matched
from .source_mapping import GROUPS
from .model import FloorplateError
from .render import Style, reference_style, html_view, powerpoint
from .sample import demonstration, source_preview

KEY = 'live_floorplate_snapshot'


def preview_only():
    # Safe default for this unmerged internal test build. UI cannot change it.
    return os.environ.get('LANDMARK_PREVIEW_ONLY','1')!='0'


def synchronize_into(state, config, layout):
    state.pop(KEY,None)
    if preview_only():
        raise FloorplateError('Live API is disabled in the internal preview deployment')
    if config.get('project') != 'Hurstville Rise':
        raise FloorplateError('Wrong configured project')
    data = sync(config)
    matched(data,layout)
    state[KEY] = data


def combine_pdfs(files):
    if len(files)==1:
        return files[0].getvalue()
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
        group = st.selectbox('Floor group',range(5),format_func=lambda g:GROUPS[g][1],key='live_floorplate_group')
        pairs = [(a,p) for a,p in matched(data,layout) if p.group==group]
        html = html_view(data,layout,pdf,group,style)
        # Hide the iframe at expiry even while the surrounding app is idle.
        remaining = int((data.synced_at+timedelta(seconds=300)-datetime.now(timezone.utc)).total_seconds()*1000)
        html = html.replace('</body>',f'<script>setTimeout(()=>{{document.body.textContent="Snapshot expired. Synchronize again."}},Math.max(0,{remaining}));</script></body>')
        if hasattr(st,'iframe'):
            st.iframe(html,height=720)
        else:
            components.html(html,height=720,scrolling=True)
        st.caption(f'{len(pairs)} Available apartments in this group. Areas shown as Internal + External.')
        st.download_button('Download five-slide as-of PowerPoint',
                           powerpoint(data,layout,pdf,style),file_name='Hurstville_Rise_Floorplates.pptx',
                           mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                           key='live_floorplate_download')
    except Exception:
        st.error('Unable to render a verified floorplate. Check every Available Lot mapping, PDF and label style. Prices are hidden.')


def render_connection_test():
    with st.expander('Arcanite connection diagnostics'):
        st.caption('Optional read-only server request. Does not synchronize prices or certify authentication, visibility or completeness. Use only in an approved, access-restricted deployment.')
        st.caption('Compare the project list and project detail independently. The guide limits the list to projects marked Show on Website; HTTP 200 alone does not prove API authentication or project visibility.')
        if st.button('Compare Project List Requests',key='compare_project_list_requests'):
            try:
                comparison = compare_project_list_requests(st.secrets['arcanite'])
            except Exception:
                comparison = {'A':None,'B':None}
            for variant,description in (('A','Required parameters only'),
                                        ('B','Official page 4 example: all optional parameters empty')):
                status = comparison[variant]
                st.write(f'Test {variant} — {description} — HTTP status: '
                         + (str(status) if status is not None else 'unavailable'))
            st.warning('Two read-only requests. Status comparison only; no raw responses shown. HTTP 200 does not prove authentication or completeness. Live price synchronization remains disabled in preview mode.')
        for label,button_key,endpoint in (
                ('Test Arcanite Connection','test_arcanite_connection','getProject'),
                ('Test Project List Access','test_arcanite_project_list','getProjectList')):
            if st.button(label,key=button_key):
                try:
                    report = connection_report(st.secrets['arcanite'],endpoint=endpoint)
                except Exception:
                    report = {'http_status':None,'ok':False}
                st.caption('Endpoint: ' + endpoint)
                status = report['http_status']
                st.write('HTTP status: ' + (str(status) if status is not None else 'unavailable'))
                if not report['ok']:
                    st.error('Connection or response validation failed. Check server-side configuration and authorized access. Live price synchronization remains disabled.')
                summary = report.get('summary')
                if summary:
                    st.dataframe([{'Field':key,'Types':', '.join(types)}
                                  for key,types in summary['response_fields'].items()],hide_index=True)
                    for array in summary['property_arrays']:
                        st.caption('Property array field: ' + array['rows_path'])
                        st.write({key:array[key] for key in ('row_count','unique_lots','invalid_lot_rows','duplicate_lot_rows')})
                        st.dataframe([{'Field':key,'Types':', '.join(types)}
                                      for key,types in array['fields'].items()],hide_index=True)
                        st.dataframe([{'Status':key,'Count':value}
                                      for key,value in array['status_counts'].items()],hide_index=True)
                    st.warning('Structure diagnostics only. Live prices, full inventory visibility and completeness remain unverified; no inventory snapshot was created.')


def render_page():
    st.header('Live Floorplate — Hurstville Rise')
    render_connection_test()
    modes=['Source PDF preview (synthetic inventory)','Synthetic demonstration']
    if not preview_only():
        modes.append('Configured Agent Portal API')
    else:
        st.info('Internal test deployment — live price synchronization disabled. Synthetic preview needs no credentials; optional connection diagnostics use server-side Secrets only.')
    if st.session_state.get('live_floorplate_mode') not in modes:
        st.session_state.pop('live_floorplate_mode',None)
    mode = st.radio('Data source',modes,key='live_floorplate_mode')
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
        mapping = st.file_uploader('Optional replacement source-checked mapping JSON',type=['json'],key='live_floorplate_mapping')
        reference = st.file_uploader('Reference Landmark PowerPoint for label styling',type=['pptx'],key='live_floorplate_style')
        clean = st.checkbox('PDF backgrounds contain no old prices; I have reviewed the source-checked mapping for this run.',key='live_floorplate_clean')
        context = hashlib.sha256(b''.join(f.getvalue() for f in originals)+(mapping.getvalue() if mapping else b'')+
                                 (reference.getvalue() if reference else b'')+str(clean).encode()).hexdigest()
        if st.session_state.get('live_floorplate_context') != context:
            st.session_state.pop(KEY,None)
            st.session_state['live_floorplate_context'] = context
        if not (originals and reference and clean):
            st.info('Provide the original PDFs, confirmed mapping and reference styling before synchronizing. No inferred positions are used.')
            return
        try:
            pdf = combine_pdfs(originals)
            with fitz.open(stream=pdf,filetype='pdf') as doc:
                mapping_data=json.loads(mapping.getvalue()) if mapping else json.loads(Path(__file__).with_name('hurstville_mapping.json').read_text())
                if mapping_data.get('version')!=2:
                    raise FloorplateError('Source-evidenced mapping required for real artwork')
                layout = parse_layout(mapping_data,len(doc),pdf)
            style = reference_style(reference.getvalue())
        except Exception:
            st.session_state.pop(KEY,None)
            st.error('Floorplate assets or mapping failed validation. No prices will be displayed.')
            return
        if mode=='Source PDF preview (synthetic inventory)':
            st.warning('SOURCE ARTWORK WITH SYNTHETIC INVENTORY — all 115 Lots are marked Available only for layout review. Unit names, bedrooms, bathrooms, areas and prices are test values, not Arcanite data.')
            refresh=st.button('Refresh synthetic source-layout preview')
            if KEY not in st.session_state or refresh:
                st.session_state[KEY]=source_preview(layout)
        elif st.button('Synchronize authorized Agent Portal inventory',type='primary'):
            # Invalidate BEFORE networking: failed refresh cannot retain old prices.
            st.session_state.pop(KEY,None)
            try:
                config = dict(st.secrets['arcanite'])
                synchronize_into(st.session_state,config,layout)
            except Exception:
                # Do not expose URLs, headers, upstream errors or response bodies.
                st.error('Synchronization failed or completeness/access could not be verified. Prices and export are unavailable. Ask the authorized operator to verify the contract.')
    show_snapshot(layout,pdf,style)
