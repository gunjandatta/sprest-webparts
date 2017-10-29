# SharePoint React WebParts
This project provides examples of using the [gd-sprest](https://gunjandatta.github.io/sprest/) and [gd-sprest-react](https://github.com/gunjandatta/sprest-react) libraries. These solutions are designed for SharePoint 2013, but is supported in SharePoint 2013/2016 and Office 365.

### Blog Post
Refer to this [blog post](http://dattabase.com/sharepoint-react-webparts/) for details on the available webparts to inherit from. The blog post gives an overview of the code for the "Document View WebPart" example.

### Batch Demo WebPart
An example of the executing batch operations. Refer to this [link](https://social.technet.microsoft.com/wiki/contents/articles/30044.sharepoint-online-performing-batch-operations-using-rest-api.aspx) for Microsoft's documentation for performing batch operations using the SharePoint Online REST API.

### Document View WebPart
Simple example of using the WebPartSearch classes to display documents. The search webpart class allows the user to select the searchable fields, currently limited to:
* Choice/Multi-Choice
* Lookup/Multi-Lookup
* Taxonomy
* Text

A mapper will be generated from the field values, and used as a tag picker for filtering the documents. Clicking on an office document will display it in the office app.

### Email WebPart
Simple example of using the SP People Picker component and the email class of the gd-sprest library.

### Hello World WebPart
Probably the first one you should look at. It's a simple example of using the webpart component, and the OnRenderDisplay and OnRenderEdit events. The page will render the webpart id in display mode, and a message indicating it's an edit mode when the page is being edited.

### List WebPart
An example of extending the WebPart Search class to display a list view and item form. The "ItemForm" component will be used to render the item form panel.

### WebPart Tabs
An example of using the WebPart Tabs class to generate a tab for each associated webpart in the same zone.