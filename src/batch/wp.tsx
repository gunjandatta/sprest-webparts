import * as React from "react";
import { List, SPTypes, Types, Web } from "gd-sprest";
import { PrimaryButton, Spinner } from "office-ui-fabric-react";

/**
 * State
 */
interface State {
    executeFl: boolean;
    list: Types.IListResult;
}

/**
 * Batch WebPart
 */
export class BatchWebPart extends React.Component<null, State> {
    /**
     * Constructor
     */
    constructor(props) {
        super(props);

        // Set the state
        this.state = {
            executeFl: false,
            list: null
        };
    }

    // Render the webpart
    render() {
        // See if the list exist
        if (this.state.list == null) {
            // Load the list
            this.loadList();

            // Show a loading message
            return (
                <Spinner label="Loading the list data..." />
            );
        }

        // See if we are executing anything
        if (this.state.executeFl) {
            // Show an execution message
            return (
                <Spinner label="Executing the request..." />
            );
        }

        // See if the list exists
        if (this.state.list.existsFl) {
            // Render a message
            return (
                <div>
                    <div>{"This list contains " + this.state.list.ItemCount + " items."}</div>
                    <PrimaryButton
                        onClick={this.createListItems}
                        text="Add 10 Items"
                    />
                    <PrimaryButton
                        onClick={this.deleteList}
                        text="Delete List"
                    />
                </div>
            );
        }

        // Render a button to create the list
        return (
            <PrimaryButton
                onClick={this.createList}
                text="Create List"
            />
        )
    }

    /**
     * Methods
     */

    // Method to create the list
    private createList = (ev: React.MouseEvent<HTMLButtonElement>) => {
        // Prevent postback
        ev.preventDefault();

        // Update the state
        this.setState({ executeFl: true });

        // Get the lists
        (new Web()).Lists()
            // Add the list
            .add({
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                Title: "DemoBatch"
            })
            // Execute the request
            .execute(list => {
                // Update the state
                this.setState({
                    executeFl: false,
                    list
                });
            });
    }

    // Method to create the list items
    private createListItems = (ev: React.MouseEvent<HTMLButtonElement>) => {
        // Prevent postback
        ev.preventDefault();

        // Update the state
        this.setState({ executeFl: true })

        // Get the current web
        let web = new Web();

        // Loop 10 times
        let ctr = 0;
        do {
            // Get the list
            web.Lists("DemoBatch")
                // Get the items
                .Items()
                // Add an item
                .add({
                    __metadata: {
                        type: "SP.Data.DemoBatchListItem"
                    },
                    Title: "Batch Item " + (++ctr)
                })
                // Batch the new items as one request
                .batch(ctr > 1);
        } while (ctr < 1);

        // Get the list
        web.Lists("DemoBatch").batch(list => {
            // Update the state
            this.setState({ executeFl: false, list });
        });

        // Execute the batch requests
        web.execute();
    }

    // Delete the list
    private deleteList = (ev: React.MouseEvent<HTMLButtonElement>) => {
        // Prevent postback
        ev.preventDefault();

        // Update the state
        this.setState({ executeFl: true }, () => {
            // Delete the list
            this.state.list.delete().execute(() => {
                // Clear the state
                this.setState({ executeFl: false, list: null });
            });
        });
    }

    // Load the list
    private loadList = () => {
        // Get the list
        (new List("DemoBatch")).execute(list => {
            // Update the state
            this.setState({ list });
        });
    }
}