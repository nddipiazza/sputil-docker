cd /sharepoint-test-utility-dotnet
./SharepointTestUtilityDotnet/build.sh
export MONO_INSTALLATION=""
if [ "$(uname)" == "Darwin" ]; then
	export MONO_INSTALLATION="/Library/Frameworks/Mono.framework/Home"
fi
./mkbundle.sh
