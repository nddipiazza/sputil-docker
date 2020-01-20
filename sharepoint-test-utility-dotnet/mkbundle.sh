mkbundle --simple --static --deps \
	-o SharepointTestUtilityDotnet/bin/Debug/sputil \
	--config $MONO_INSTALLATION/etc/mono/config \
	--machine-config $MONO_INSTALLATION/etc/mono/4.5/machine.config \
	-L SharepointTestUtilityDotnet/bin/Debug \
	SharepointTestUtilityDotnet/bin/Debug/SharepointTestUtilityDotnet.exe \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.Office.Client.Policy.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.Office.Client.TranslationServices.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.Online.SharePoint.Client.Tenant.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.ProjectServer.Client.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.DocumentManagement.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.Publishing.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.Runtime.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.Search.Applications.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.Search.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.Taxonomy.dll \
	./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.UserProfiles.dll \
./SharepointTestUtilityDotnet/bin/Debug/Microsoft.SharePoint.Client.WorkflowServices.dll