import * as React from "react";
import { Utility } from "gd-sprest";
import { Components } from "gd-sprest-react";
import { PrimaryButton, TextField, ITextField } from "office-ui-fabric-react";
declare var SP;

/**
 * Email
 */
export class EmailWebPart extends React.Component<null, null> {
    private _spPicker: Components.SPPeoplePicker = null;
    private _tb: ITextField = null;

    // Render the component
    render() {
        return (
            <div>
                <Components.SPPeoplePicker ref={picker => { this._spPicker = picker; }} />
                <TextField multiline={true} rows={6} componentRef={tb => { this._tb = tb; }} />
                <PrimaryButton text="Email" onClick={this.sendEmail} />
            </div>
        );
    }

    /**
     * Methods
     */

    // Method to send an email
    private sendEmail = (ev: React.MouseEvent<HTMLButtonElement>) => {
        // Prevent postback
        ev.preventDefault();

        // Get the selected user
        let user = this._spPicker.state.personas[0];
        if (user) {
            // Display a pop-up message
            SP.SOD.execute("sp.ui.dialog.js", "SP.UI.ModalDialog.showWaitScreenWithNoClose", "Sending Email", "Attempting to send the email. This dialog will close after the request completes.");

            // Email the user
            Utility().sendEmail({
                Body: this._tb.value,
                Subject: "Demo Email",
                To: [user.secondaryText]
            }).execute(() => {
                // Close the dialog
                SP.SOD.execute("sp.ui.dialog.js", "SP.UI.ModalDialog.commonModalDialogClose");
            });
        }
    }
}