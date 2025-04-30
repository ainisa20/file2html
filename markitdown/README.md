
docker run -d \
  --name markdown-flask \
  --restart always \
  -p 5009:5009 \
  -v $(pwd)/uploads:/app/uploads \
  -v $(pwd)/outputs:/app/outputs \
  -v /var/run/docker.sock:/var/run/docker.sock \
  file2md
