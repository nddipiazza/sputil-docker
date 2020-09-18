# sputil-docker
Dockerbuild for an SP utility


# how to publish a new version

Merge your change. Create a tag. 

run `docker build .` from the root directory. It will publish the tag to docker.

The result, you can then use the new docker image:

Example:

`ndipiazza/sputil_mono:0.0.4`



