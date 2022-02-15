FROM mono:6.12.0.122-slim

COPY sharepoint-test-utility-dotnet /sharepoint-test-utility-dotnet

RUN apt-get update \
  && apt-get install -y binutils curl mono-devel ca-certificates-mono fsharp mono-vbnc nuget referenceassemblies-pcl \
  && rm -rf /var/lib/apt/lists/* /tmp/* \
  && /sharepoint-test-utility-dotnet/build.sh