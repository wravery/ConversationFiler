import { Data } from "./Model";
import { RESTData } from "./RESTData";
import { EWSData } from "./EWSData";

export module Factory {
    // Use the RESTData provider if possible, but if it's not supported, fallback to the EWSData provider
    export function getData(mailbox: Office.Mailbox): Data.IModel {
        return mailbox.restUrl
            ? new RESTData.Model(mailbox)
            : new EWSData.Model(mailbox);
    }
}