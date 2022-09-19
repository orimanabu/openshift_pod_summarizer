#!/usr/bin/python3

import sys
import copy
import json
import yaml
import pprint
import openpyxl
import subprocess

masters = []
workers = []

def load_nodes():
    output = subprocess.run('kubectl get node -o json'.split(), capture_output=True)
    json_data = json.loads(output.stdout)

    for item in json_data['items']:
        hostname = item['metadata']['name']
        labels = item['metadata']['labels']
        if 'node-role.kubernetes.io/master' in labels:
            masters.append(hostname)
        if 'node-role.kubernetes.io/worker' in labels:
            workers.append(hostname)

def print_nodes():
    print('# masters:')
    pprint.pprint(masters)
    print('# workers:')
    pprint.pprint(workers)

def write_row(sheet, array, start_row, start_col, color):
    for x, cell in enumerate(array):
        sheet.cell(row=start_row, column=start_col + x, value=array[x])
        sheet.cell(row=start_row, column=start_col + x).fill = color

def hostname2role(hostname):
    role = ''
    if hostname in masters:
        role = role + 'master'
    if hostname in workers:
        if role != '':
            role = role + '/'
        role = role + 'worker'
    return role

def dict2json(obj):
    # return json.dumps(obj, indent=2)
    # return yaml.dump(obj, Dumper=yaml.CDumper)
    return yaml.dump(obj, Dumper=yaml.Dumper).rstrip()

def ns_pod_key(ns, pod):
    return '{}__{}'.format(ns, pod)

def normalize_pod_name(name, owner_kind, node_name):
    array = name.split('-')
    if owner_kind == 'DaemonSet' or owner_kind == 'CatalogSource':
        array[-1] = 'X' * len(array[-1])
        return '-'.join(array)
    if owner_kind == 'ReplicaSet':
        array[-2] = 'X' * len(array[-2])
        array[-1] = 'X' * len(array[-1])
        return '-'.join(array)
    if owner_kind == 'StatefulSet':
        return name[0:-1] + 'X'
    if owner_kind == 'Node':
        return name.replace(node_name, 'HOSTNAME')
    if name.startswith('etcd-guard') or name.startswith('kube-apiserver-guard') or name.startswith('kube-controller-manager-guard') or name.startswith('openshift-kube-scheduler-guard'):
        return name.replace(node_name, 'HOSTNAME')
    return name

def normalize_owner_name(owner_kind, owner_name):
    array = owner_name.split('-')
    if owner_kind == 'ReplicaSet':
        array[-1] = 'X' * len(array[-1])
        return '-'.join(array)
    if owner_kind == 'Node':
        return 'HOSTNAME'
    return owner_name

def normalize_owner_kind(owner_kind, owner_name, ns):
    if owner_kind == 'ReplicaSet':
        output = subprocess.run('kubectl -n {} get {}/{} -o json'.format(ns, owner_kind, owner_name).split(), capture_output=True)
        json_data = json.loads(output.stdout)
        refs = json_data['metadata'].get('ownerReferences')
        if refs:
            kind = refs[0]['kind']
            return '{} ({})'.format(owner_kind, kind)
    return owner_kind

def get_number_of_pods(selector, owner_kind, owner_name, ns):
    if owner_kind == 'DaemonSet':
        print('  ### selector:{}'.format(selector))
        # if 'node-role.kubernetes.io/master' in selector:
        if selector.get('node-role.kubernetes.io/master') == '':
            return '# of masters'
        # if 'node-role.kubernetes.io/worker' in selector:
        if selector.get('node-role.kubernetes.io/worker') == '':
            return '# of workers'
        if selector.get('kubernetes.io/os') == 'linux' or selector.get('beta.kubernetes.io/os') == 'linux':
            return '# of linux nodes'
        return 'unknown'

    if owner_kind == 'Node' and ns == 'openshift-etcd':
        return 'Static Pod on masters'
    if owner_kind == 'Node' and ns == 'openshift-kube-apiserver':
        return 'Static Pod on masters'
    if owner_kind == 'Node' and ns == 'openshift-kube-controller-manager':
        return 'Static Pod on masters'
    if owner_kind == 'Node' and ns == 'openshift-kube-scheduler':
        return 'Static Pod on masters'

    output = subprocess.run('kubectl -n {} get {}/{} -o json'.format(ns, owner_kind, owner_name).split(), capture_output=True)
    json_data = json.loads(output.stdout)
    replicas = json_data['spec'].get('replicas')
    return 'replicas={}'.format(replicas)

load_nodes()
print_nodes()

book = openpyxl.Workbook()
sheet = book.active
sheet.title = 'Pods'
current_row = 1
fill_header = openpyxl.styles.PatternFill(patternType='solid', fgColor='D9EAD3')
# sheet.freeze_panes = 'A2'
sheet.freeze_panes = 'C2'

output = subprocess.run('kubectl get pod -A -o json'.split(), capture_output=True)
json_data = json.loads(output.stdout)

header_labels = [
    'ns',
    'pod_name',
    'num_of_pods',
    'owner_kind',
    'owner_name',
    'affinity',
    'dnsPolicy',
    'enableServiceLinks',
    'hostNetwork',
    'hostPID',
    # 'nodeName',
    # 'role',
    'nodeSelector',
    'preemptionPolicy',
    'priority',
    'priorityClassName',
    'restartPolicy',
    'schedulerName',
    'serviceAccount',
    'serviceAccountName',
    'pod_securityContext',
    'tolerations',
    'terminationGracePeriodSeconds',
    'qosClass',
    'container_name',
    'container_image',
    'container_imagePullPolicy',
    'container_resources',
    'container_securityContext'
]

pod_exists = {}

header2column = {}
for i, label in enumerate(header_labels, 1):
    # openxl column is 1-origin
    header2column[label] = i

def xls_input_cell_by_key(sheet, row, key, value):
    sheet.cell(row=row, column=header2column[key], value=value)
    sheet.cell(row=row, column=header2column[key]).alignment = openpyxl.styles.Alignment(vertical='center')
    sheet.cell(row=row, column=header2column[key]).font = openpyxl.styles.fonts.Font(name='Source Code Pro Medium')

write_row(sheet, header_labels, 1, 1, fill_header)
current_row = current_row + 1

prev_pod_name = ''
for item in json_data['items']:
    md = item['metadata']
    spec = item['spec']
    status = item['status']
    row = []

    if status['phase'] != 'Running':
        continue

    print('* ns:{}, pod_name:{}, phase:{}'.format(md['namespace'], md['name'], status['phase']))
    xls_input_cell_by_key(sheet, current_row, 'ns', md['namespace'])
    xls_input_cell_by_key(sheet, current_row, 'pod_name', md['name'])

    refs = md.get('ownerReferences')
    if refs:
        ref = refs[0]
        print('  owner_kind:{}, owner_name:{}'.format(ref['kind'], ref['name']))

        # rename 'pod_name' column
        pod_name = normalize_pod_name(md['name'], ref['kind'], spec['nodeName'])
        key = ns_pod_key(md['namespace'], pod_name)
        if pod_exists.get(key):
            print('  => SKIP')
            continue
        pod_exists[key] = True
        xls_input_cell_by_key(sheet, current_row, 'pod_name', pod_name)

        xls_input_cell_by_key(sheet, current_row, 'owner_kind', normalize_owner_kind(ref['kind'], ref['name'], md['namespace']))
        xls_input_cell_by_key(sheet, current_row, 'owner_name', normalize_owner_name(ref['kind'], ref['name']))

        # 'number of pods' column
        xls_input_cell_by_key(sheet, current_row, 'num_of_pods', get_number_of_pods(spec.get('nodeSelector', ''), ref['kind'], ref['name'], md['namespace']))

        if len(refs) > 1:
            print('!! more than one ownerReferences !!')
            sys.exit(1)

    else:
        pod_name = normalize_pod_name(md['name'], '', spec['nodeName'])
        key = ns_pod_key(md['namespace'], md['name'])
        if pod_exists.get(key):
            print('  => SKIP')
            continue
        pod_exists[key] = True
        xls_input_cell_by_key(sheet, current_row, 'pod_name', pod_name)

    print('  affinity:{}'.format(spec.get('affinity', '')))
    xls_input_cell_by_key(sheet, current_row, 'affinity', dict2json(spec.get('affinity', '')))
    print('  dnsPolicy:{}, hostNetwork:{}, hostPID:{}'.format(spec.get('dnsPolicy', ''), spec.get('hostNetwork', ''), spec.get('hostPID', '')))
    xls_input_cell_by_key(sheet, current_row, 'dnsPolicy', spec.get('dnsPolicy', ''))
    xls_input_cell_by_key(sheet, current_row, 'enableServiceLinks', spec.get('enableServiceLinks', ''))
    xls_input_cell_by_key(sheet, current_row, 'hostNetwork', spec.get('hostNetwork', ''))
    xls_input_cell_by_key(sheet, current_row, 'hostPID', spec.get('hostPID', ''))
    print('  nodeName:{}, role:{}, nodeSelector:{}'.format(spec.get('nodeName', ''), hostname2role(spec.get('nodeName', '')), spec.get('nodeSelector', '')))
    # xls_input_cell_by_key(sheet, current_row, 'nodeName', spec.get('nodeName', ''))
    # xls_input_cell_by_key(sheet, current_row, 'role', hostname2role(spec.get('nodeName', '')))
    xls_input_cell_by_key(sheet, current_row, 'nodeSelector', dict2json(spec.get('nodeSelector', '')))
    print('  preemptionPolicy:{}, priorityClassName:{}'.format(spec.get('preemptionPolicy', ''), spec.get('priorityClassName', '')))
    xls_input_cell_by_key(sheet, current_row, 'preemptionPolicy', spec.get('preemptionPolicy', ''))
    xls_input_cell_by_key(sheet, current_row, 'priority', spec.get('priority', ''))
    xls_input_cell_by_key(sheet, current_row, 'priorityClassName', spec.get('priorityClassName', ''))
    xls_input_cell_by_key(sheet, current_row, 'restartPolicy', spec.get('restartPolicy', ''))
    xls_input_cell_by_key(sheet, current_row, 'schedulerName', spec.get('schedulerName', ''))
    xls_input_cell_by_key(sheet, current_row, 'serviceAccount', spec.get('serviceAccount', ''))
    xls_input_cell_by_key(sheet, current_row, 'serviceAccountName', spec.get('serviceAccountName', ''))
    print('  pod_securityContext:{}'.format(spec.get('securityContext', '')))
    xls_input_cell_by_key(sheet, current_row, 'pod_securityContext', dict2json(spec.get('securityContext', '')))
    print('  tolerations:{}'.format(spec.get('tolerations', '')))
    xls_input_cell_by_key(sheet, current_row, 'tolerations', dict2json(spec.get('tolerations', '')))
    xls_input_cell_by_key(sheet, current_row, 'terminationGracePeriodSeconds', spec.get('terminationGracePeriodSeconds', ''))
    print('  qosClass:{}'.format(status.get('qosClass', '')))
    xls_input_cell_by_key(sheet, current_row, 'qosClass', status.get('qosClass', ''))

    nctrs = 0
    row_pod_container_start = 0
    for ctr in spec['containers']:
        xls_input_cell_by_key(sheet, current_row, 'container_name', ctr.get('name', ''))
        xls_input_cell_by_key(sheet, current_row, 'container_image', ctr.get('image', ''))
        xls_input_cell_by_key(sheet, current_row, 'container_imagePullPolicy', ctr.get('imagePullPolicy', ''))
        xls_input_cell_by_key(sheet, current_row, 'container_resources', dict2json(ctr.get('resources', '')))
        xls_input_cell_by_key(sheet, current_row, 'container_securityContext', dict2json(ctr.get('securityContext', '')))
        if nctrs == 0:
            print('    ctr_name:{}, ctr_resources:{}, ctr_securityContext:{}'.format(
                ctr.get('name', ''),
                ctr.get('resources', ''),
                ctr.get('securityContext', '')
            ))
            row_pod_container_start = current_row
        else:
            for col in range(header2column['ns'], header2column['qosClass'] + 1):
                sheet.merge_cells(start_row=row_pod_container_start, end_row=current_row, start_column=col, end_column=col)

        nctrs += 1
        current_row = current_row + 1

def get_cell_length(value):
    # print('*** get_cell_length(): value={}'.format(value))
    if str(value).find('\n'):
        sentences = str(value).split('\n')
        # print('*** sentences={}'.format(sentences))
        max_length = 0
        for s in sentences:
            if len(s) > max_length:
                max_length = len(s)
        return max_length
    return len(str(value))

for col in sheet.columns:
    max_length = 0
    colname = col[0].column_letter

    for cell in col:
        length = get_cell_length(cell.value)
        if length > max_length:
               max_length = length
    target_width = (max_length + 2) * 1.2
    sheet.column_dimensions[colname].width = target_width

book.save('result.xlsx')
