# YartuDocThumbnail
Convert office files to pdf or photo. using LibreOffice Uno
# Usage
```
sudo apt-get install -y libreoffice libreoffice-script-provider-python uno-libs3 python3-uno python3
```
Running as a service and listening on port 2002
```
$ soffice "-accept=socket,port=2002;urp;"
```

```
thumblier = YartuDocThumb()
thumblier.doc_to_img(file = "abc.docx", out_path = "/home/")
thumblier.doc_to_pdf(file = "abc.docx", out_path = "/home/")
```
