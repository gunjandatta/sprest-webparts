import * as React from "react";
import { SPTypes, Types } from "gd-sprest";
import { ItemForm, Panel, WebPartSearch, IWebPartSearchProps, IWebPartSearchState } from "gd-sprest-react";
import { PrimaryButton } from "office-ui-fabric-react";

/**
 * List Item Information
 */
export interface IListItem extends Types.IListItemQueryResult {
    Attachments?: boolean;
    TestBoolean?: boolean;
    TestChoice?: string;
    TestDate?: string;
    TestDateTime?: string;
    TestLookup?: Types.ComplexTypes.FieldLookupValue;
    TestLookupId?: string | number;
    TestMultiChoice?: string;
    TestMultiLookup?: string;
    TestMultiLookupId?: string;
    TestMultiUser?: { results: Array<number> };
    TestMultiUserId?: Array<number>;
    TestNote?: string;
    TestNumberDecimal?: number;
    TestNumberInteger?: number;
    TestUrl?: string;
    TestUser?: Types.ComplexTypes.FieldUserValue;
    TestUserId?: string | number;
    Title?: string;
}

/**
 * State
 */
interface State extends IWebPartSearchState {
    controlMode?: number;
    errorMessage?: string;
    item?: IListItem;
}

/**
 * List WebPart
 */
export class ListWebpart extends WebPartSearch<IWebPartSearchProps, State> {
    private _itemForm: ItemForm = null;
    private _panel: Panel = null;

    /**
     * Constructor
     */
    constructor(props) {
        super(props);

        // Update the query
        this._query.Expand = ["AttachmentFiles", "TestLookup", "TestMultiLookup", "TestMultiUser", "TestUser"];
        this._query.OrderBy = ["Title"];
        this._query.Select = ["*", "Attachments", "AttachmentFiles", "TestLookup/ID", "TestLookup/Title", "TestMultiLookup/ID", "TestMultiLookup/Title", "TestMultiUser/ID", "TestMultiUser/Title", "TestUser/ID", "TestUser/Title"];
    }

    // The render container event
    onRenderContainer = (items) => {
        let elItems = [];

        // Ensure items exist
        if (items && items.length > 0) {
            // Parse the items
            for (let i = 0; i < items.length; i++) {
                // Add the item
                elItems.push(this.onRenderItem(items[i]));
            }

            // Return the container
            return (
                <div className="ms-Grid">
                    {elItems}
                    <div className="ms-Grid-row" key="item_form">
                        <Panel headerText="Item Form" ref={panel => { this._panel = panel; }}>
                            <div className="">{this.state.errorMessage + ""}</div>
                            <ItemForm
                                controlMode={this.state.controlMode}
                                fields={[
                                    { name: "Attachments" },
                                    { name: "Title" },
                                    { name: "TestBoolean" },
                                    { name: "TestChoice" },
                                    { name: "TestDate" },
                                    { name: "TestDateTime" },
                                    { name: "TestLookup" },
                                    { name: "TestManagedMetadata" },
                                    { name: "TestMultiChoice" },
                                    { name: "TestMultiLookup" },
                                    { name: "TestMultiUser" },
                                    { name: "TestNote" },
                                    { name: "TestNumberDecimal" },
                                    { name: "TestNumberInteger" },
                                    { name: "TestUrl" },
                                    { name: "TestUser" }
                                ]}
                                item={this.state.item}
                                listName={this.props.cfg.ListName}
                            />
                        </Panel>
                    </div>
                </div>
            );
        }

        // Not items exist
        return (
            <div className="ms-fontSize-l">No items exist...</div>
        );
    }

    // The render item event
    onRenderItem = (item: IListItem) => {
        // Return the item
        return (
            <div className="ms-Grid-row" key={"item_" + item.Id}>
                <div className="ms-Grid-col ms-md-1">
                    <PrimaryButton text="View" data-itemId={item.Id} onClick={this.viewItem} />
                </div>
                <div className="ms-Grid-col ms-md-1">
                    <PrimaryButton text="Edit" data-itemId={item.Id} onClick={this.editItem} />
                </div>
                <div className="ms-Grid-col ms-md-2">
                    {item.Title ? item.Title : ""}
                </div>
                <div className="ms-Grid-col ms-md-2">
                    {item.TestChoice ? item.TestChoice : ""}
                </div>
                <div className="ms-Grid-col ms-md-2">
                    {item.TestDate ? item.TestDate : ""}
                </div>
                <div className="ms-Grid-col ms-md-2">
                    {item.TestUrl ? item.TestUrl : ""}
                </div>
                <div className="ms-Grid-col ms-md-2">
                    {item.TestUser ? item.TestUser.Title : ""}
                </div>
            </div>
        );
    }

    /**
     * Methods
     */

    // Method to edit the item
    private editItem = (el: React.MouseEvent<HTMLButtonElement>) => {
        // Prevent postback
        el.preventDefault();

        // Clear the selected item
        this.setState({
            controlMode: SPTypes.ControlMode.Edit,
            item: this.getItem(parseInt(el.currentTarget.getAttribute("data-itemId")))
        }, () => {
            // Show the panel
            this._panel.show();
        });
    }

    // Method to get the item
    private getItem = (itemId) => {
        // Parse the items
        for (let i = 0; i < this.state.items.length; i++) {
            let item = this.state.items[i];

            // See if this is the target item
            if (itemId == item.Id) {
                // Return the item
                return item;
            }
        }

        // Item not found
        return null;
    }

    // Method to view the item
    private viewItem = (el: React.MouseEvent<HTMLButtonElement>) => {
        // Prevent postback
        el.preventDefault();

        // Clear the selected item
        this.setState({
            controlMode: SPTypes.ControlMode.Display,
            item: this.getItem(parseInt(el.currentTarget.getAttribute("data-itemId")))
        }, () => {
            // Show the panel
            this._panel.show();
        });
    }
}