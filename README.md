# redmine_preview_office

Plugin for Redmine. Show Microsoft Office attachments in preview pane. 
Supports .doc, .docx, .xls, .xlsx, .ppt, .pptx, .rtf, .odt

![PNG that represents a quick overview](/doc/Overview.png)

### Use case(s)

* View contents of a Microsoft Office file as pdf in Redmine's preview pane 

### Install

1. Install [libreoffice]:https://www.libreoffice.org (many linux distributions provide ready installers) 

2. download plugin and copy plugin folder redmine_preview_office go to Redmine's plugins folder 

3. restart server f.i.  

`sudo /etc/init.d/apache2 restart`

### Uninstall

1. go to plugins folder, delete plugin folder redmine_preview_office

`rm -r redmine_preview_office`

2. restart server f.i. 

`sudo /etc/init.d/apache2 restart`

### Use

* Go to Administration->Information and verify if Libreoffice is available
* On Issue show page click on a docx or other Microsoft Office attachment and view its contents in Redmine's preview pane as a pdf file

**Have fun!**

### Localisations

* 1.0.1 
  - English
  - German
* 1.0.0 - no localizable strings present in plugin

### Change-Log* 

**1.0.3** 
  - cleaned code
  - added support for MS Windows platform

**1.0.2** a
  - dded option to embed preview via `<object><embed>`-tag or via `<iframe>`-tag

**1.0.1** 
  - added check for libreoffice

**1.0.0** 
  - running on Redmine 3.4.6 
