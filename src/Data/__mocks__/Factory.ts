import { Data } from "../Model";

export module Factory {
    export interface MockMailbox extends Office.Mailbox {
        mockResults: Data.Match[];
    }

    class MockData implements Data.IModel {
        constructor(private mockResults: Data.Match[]) {
        }

        getItemsAsync(onLoadComplete: (results: Data.Match[]) => void, onProgress: (progress: Data.Progress) => void, onError: (message: string) => void): void {
            if (this.mockResults) {
                if (this.mockResults.length > 0) {
                    onLoadComplete(this.mockResults);
                } else {
                    onError("Empty results");
                }
            } else {
                onProgress(Data.Progress.GetCallbackToken);
            }
        }

        moveItemsAsync(folderId: string, onMoveComplete: (count: number) => void, onError: (message: string) => void): void {
            onMoveComplete(this.mockResults.length);
        }
    }

    // Use the MockData provider
    export function getData(mailbox: Office.Mailbox): Data.IModel {
        return new MockData((<MockMailbox>mailbox).mockResults);
    }
}