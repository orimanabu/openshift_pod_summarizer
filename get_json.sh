#!/bin/bash

oc whoami > /dev/null 2>&1
if [ x"$?" != x"0" ]; then
	echo "needs oc login, exit"
	exit 1
fi

# kubectl get -A pod,node,replicaset,statefulset,deployment,catalogsource,job,cronjob -o json > all.json
# kubectl get -A pod,replicaset,statefulset,deployment,catalogsource,job,cronjob -o json > all.json
kubectl get -A pod,replicaset,statefulset,deployment,catalogsource,job,cronjob -o yaml > all.yaml
