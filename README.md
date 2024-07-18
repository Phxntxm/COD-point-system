# Running

This repo provides a Dockerfile that can be used to simplify running this script, as long as you have docker installed you should be able to build and run this pretty simply. I'll provide the linux instructions here, as I don't know Windows well at all, so unsure how to properly handle docker on windows.

```bash
# Builds a docker container with the name of codpoints
docker build -t codpoints:latest .
# Runs the docker container, providing a volume so that they file output can be saved locally
docker run -v ${PWD}/codpointsoutput/:/app/codpointsoutput/ --name=codpointssystem codpoints:latest
# Once this run finishes, you should have a directory in the current directory called codpointsoutput
#  with a file in it called output.xlsx. That is the finished Excel file for this
```
