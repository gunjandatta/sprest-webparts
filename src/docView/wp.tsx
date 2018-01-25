import * as React from "react";
import { ContextInfo } from "gd-sprest";
import { WebParts } from "gd-sprest-react";
import { Icon } from "office-ui-fabric-react";
declare var SP;

/**
 * Document Item
 */
interface IDocument extends WebParts.IWebPartSearchItem {
    DocIcon: string;
    FileRef: string;
    LinkFilename: string;
}

/**
 * Document View
 */
export class DocViewWebPart extends WebParts.WebPartSearch {
    /**
     * Constructor
     */
    constructor(props) {
        super(props);

        // Enable caching
        this._cacheFl = true;

        // Update the query to order by the filename, and include the document fields
        this._query.GetAllItems = true;
        this._query.OrderBy = ["LinkFilename"];
        this._query.Select = ["DocIcon", "FileRef", "ID", "LinkFilename"];
        this._query.Top = 500;
    }

    /**
     * Events
     */

    // The doc icon click event
    onDocIconClicked = (ev: React.MouseEvent<HTMLDivElement>) => {
        // Prevent postback
        ev.preventDefault();

        // Get the document url
        let docUrl = ev.currentTarget.getAttribute("data-docurl");

        // See if this is an office document
        if (ev.currentTarget.getAttribute("data-isofficedoc") == "true") {
            // Display the document in a modal dialog
            SP.SOD.execute("sp.ui.dialog.js", "SP.UI.ModalDialog.showModalDialog", {
                showMaximized: true,
                title: "",
                url: docUrl
            });
        } else {
            // Open the document in a new window/tab
            window.open(docUrl, "_blank");
        }
    }

    // The render container event
    onRenderContainer = (items: Array<IDocument>) => {
        let elements = [];

        // Parse the items
        for (let i = 0; i < items.length; i++) {
            // Add the item
            elements.push(this.onRenderItem(items[i]));
        }

        // Render the container
        return (
            <div className="docView">
                <div className="docView-row">{elements}</div>
            </div>
        );
    }

    // The render item event
    onRenderItem = (item: IDocument) => {
        let isOfficeDocFl = false;

        // Compute the document image url
        let docUrl = ContextInfo.webAbsoluteUrl + "/_layouts/15/WopiFrame2.aspx?sourcedoc=" + item.FileRef + "&action=present";

        // Determine the icon name
        let iconName = null;
        switch (item.DocIcon) {
            case "docx":
                iconName = "WordLogo";
                isOfficeDocFl = true;
                break;
            case "pdf":
                iconName = "PDF";
                break;
            case "pptx":
                iconName = "PowerPointLogo";
                isOfficeDocFl = true;
                break;
            case "vsdx":
                iconName = "VisioLogo";
                isOfficeDocFl = true;
                break;
            case "xlsx":
                iconName = "ExcelLogo";
                isOfficeDocFl = true;
                break;
            default:
                iconName = "Document";
                break;
        }

        // Render the item
        return (
            <div className="docView-item"
                data-docurl={isOfficeDocFl ? docUrl : item.FileRef}
                data-isofficedoc={isOfficeDocFl}
                key={"item_" + item.Id}
                onClick={this.onDocIconClicked}>
                <Icon className="docView-icon" iconName={iconName} />
                <span className="dovView-title">{item.LinkFilename}</span>
            </div>
        );
    }
}