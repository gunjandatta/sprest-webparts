import * as React from "react";
import { List, SPTypes, Types, Web } from "gd-sprest";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Spinner } from "office-ui-fabric-react/lib/Spinner";

/**
 * State
 */
interface State {
    executeFl: boolean;
    list: Types.SP.IListResult;
    loadFl: boolean;
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
            list: null,
            loadFl: false
        };
    }

    // Render the webpart
    render() {
        // See if the list exist
        if (this.state.list == null && !this.state.loadFl) {
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
        if (this.state.list) {
            // Render a message
            return (
                <div>
                    <div>{"This list contains " + this.state.list.ItemCount + " items."}</div>
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
        this.setState({ executeFl: true })

        // Get the current web
        let web = Web();

        // Get the lists
        web.Lists()
            // Create the list
            .add({
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                Title: "DemoBatch"
            })
            // Batch the request
            .batch();

        // Loop 10 times
        let ctr = 0;
        do {
            // Get the list
            web.Lists("DemoBatch")
                // Get the items
                .Items()
                // Add an item
                .add({
                    Title: "Batch Item " + (++ctr)
                })
                // Batch the new items as one request
                .batch(ctr > 1);
        } while (ctr < 10);

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
        List("DemoBatch").execute(
            // Success
            list => {
                // Update the state
                this.setState({
                    list,
                    loadFl: true
                });
            },
            // Error
            () => {
                // Update the state
                this.setState({
                    loadFl: true
                });
            }
        );
    }
}