import * as React from "react";
import { ContextInfo } from "gd-sprest";
import { WebPartSearch, IWebPartSearchItem } from "gd-sprest-react";
declare var SP;

/**
 * Document Item
 */
interface IDocument extends IWebPartSearchItem {
    DocIcon: string;
    FileRef: string;
    LinkFilename: string;
}

/**
 * Document View
 */
export class DocViewWebPart extends WebPartSearch {
    /**
     * Constructor
     */
    constructor(props) {
        super(props);

        // Update the query to order by the filename, and include the document fields
        this._query.OrderBy = ["LinkFilename"];
        this._query.Select = ["DocIcon", "FileRef", "ID", "LinkFilename"];
    }

    /**
     * Events
     */

    // The doc icon click event
    onDocIconClicked = (ev: React.MouseEvent<HTMLDivElement>) => {
        // Prevent postback
        ev.preventDefault();

        // Get the document url
        let docUrl = ev.currentTarget.getAttribute("data-docUrl");

        // See if this is an office document
        if (ev.currentTarget.getAttribute("data-isOfficeDoc") == "true") {
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
            <div className="ms-Grid">
                <div className="ms-Grid-row">{elements}</div>
            </div>
        );
    }

    // The render item event
    onRenderItem = (item: IDocument) => {
        let isOfficeDocFl = false;

        // Compute the document image url
        let docUrl = ContextInfo.webAbsoluteUrl + "/_layouts/15/WopiFrame2.aspx?sourcedoc=" + item.FileRef + "&action=present";

        // Determine the icon
        let icon = "";
        switch (item.DocIcon) {
            case "docx":
                icon = "WordLogo";
                isOfficeDocFl = true;
                break;
            case "pdf":
                icon = "PDF";
                break;
            case "pptx":
                icon = "PowerPointLogo";
                isOfficeDocFl = true;
                break;
            case "vsdx":
                icon = "VisioLogo";
                isOfficeDocFl = true;
                break;
            case "xlsx":
                icon = "ExcelLogo";
                isOfficeDocFl = true;
                break;
            default:
                icon = "Document";
                break;
        }

        // Render the item
        return (
            <div
                className="ms-Grid-col ms-md1 ms-textAlignCenter docView-item"
                data-docUrl={isOfficeDocFl ? docUrl : item.FileRef}
                data-isOfficeDoc={isOfficeDocFl}
                key={"item_" + item.Id}
                onClick={this.onDocIconClicked}>
                <i className={"ms-fontSize-su ms-Icon ms-Icon--" + icon} />
                <span className="ms-fontSize-mPlus">{item.LinkFilename}</span>
            </div>
        );
    }
}