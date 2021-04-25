import urllib

url="https://www1.sii.cl/cgi-bin/Portal001/mipeShowPdf.cgi?CODIGO=1241390875"
urllib.request.urlretrieve(url, filename)