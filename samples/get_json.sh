#!/bin/bash

oc whoami > /dev/null 2>&1
if [ x"$?" != x"0" ]; then
	echo "needs oc login, exit"
	exit 1
fi

# oc get pod -A -o json > pods.json
# oc get node -A -o json > nodes.json
# oc get replicaset -A -o json > replicasets.json
# oc get deployment -A -o json > deployments.json
# oc get statefulset -A -o json > statefulsets.json
# oc get daemonset -A -o json > daemonsets.json
# oc get catalogsource -A -o json > catalogsources.json
kubectl get -A pod,node,replicaset,statefulset,deployment,catalogsource -o json > all.json
