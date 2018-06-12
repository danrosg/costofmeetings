
#!/usr/bin/env bash

az webapp deployment user set --user-name costofmeetings --password Pass@@%%202
az group create --name hackOptimizersRG --location "eastus"
az appservice plan create --name hackOptimizersASP --resource-group hackOptimizersRG --sku S1 --is-linux
az webapp create --resource-group hackOptimizersRG --plan hackOptimizersASP --name hackOptimizers --runtime "NODE|8.11" --deployment-local-git --startup-file "npm run run_azure"


#https://costofmeetings@hackoptimizers.scm.azurewebsites.net/hackOptimizers.git


git remote add azure https://costofmeetings@hackoptimizers.scm.azurewebsites.net/hackOptimizers.git

git push azure master
