'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
Browser("住友").Sync
Browser("住友").CloseAllTabs