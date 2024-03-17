// This is created as example by Qianfu Ni

// eslint-disable-next-line @typescript-eslint/no-unused-vars
module Smartwise.Crm {
    export class ContactEntityForm {
        static xrmApiUri = 'api/data/v9.1';
        static BatchQueryContentType = 'multipart/mixed;boundary=batch_fetchquery';
        static JsonContentType = 'application/json; charset=utf-8';

        protected formContext: Xrm.FormContext;
        protected executionContext: Xrm.Events.EventContext;
        protected globalContext: Xrm.GlobalContext;
        protected clientUrl: string = null;
        protected xrmApiUrl: string;
        protected xrmApiBatchUrl: string;

        onLoad(executionContext: any): void {
            this.executionContext = executionContext;
            this.formContext = executionContext.getFormContext();
            this.globalContext = typeof GetGlobalContext !== 'undefined' ? GetGlobalContext() : Xrm.Utility.getGlobalContext();
            this.clientUrl = this.globalContext.getClientUrl();
            this.xrmApiUrl = `${this.clientUrl}/${ContactEntityForm.xrmApiUri}/`;
            this.xrmApiBatchUrl = `${this.xrmApiUrl}$batch`;
            this.addOnChangeEventHandler('firstname', this.onFirstNameChange.bind(this));
        }

        addOnChangeEventHandler(attributeName: string, eventHandler: any): void {
            const attributeControl = this.formContext.getAttribute(attributeName);
            if (attributeControl == null) {
                console.warn(`${attributeName} is not on the current form. form entity name: ${this.formContext.ui.setFormEntityName}`);
                return;
            }
            attributeControl.addOnChange(eventHandler);
        }

        onFirstNameChange(_: Xrm.Events.EventContext): void {
            var fetchXml = [
                '<fetch top="10">',
                '  <entity name="contact">',
                '    <attribute name="fullname"/>',
                '  </entity>',
                '</fetch>'
                ].join('');

            this.fetchEntities('contacts', fetchXml, (result: ProcessResult<Entity[]>): void  => {
                console.log("Contact fetch result: ", result.result);
            });

            console.debug("Contact first name is changed.")
        }

        private fetchEntities(entitySetName: string, fetchXml: string, callback: (result: ProcessResult<Entity[]>) => void): void {
            const processResult = new ProcessResult<Entity[]>();
            const body = this.getFetchQueryBody(entitySetName, fetchXml);

            let httpStatus = 0;
            fetch(this.xrmApiBatchUrl, {
                body,
                headers: this.getHeaders(ContactEntityForm.BatchQueryContentType),
                credentials: 'same-origin',
                method: 'POST'
            })
                .then((response: any) => {
                    httpStatus = response.status;
                    return response.text();
                })
                .then((data: string) => {
                    if (httpStatus !== 200 && httpStatus !== 204) {
                        processResult.errorMessage = data;
                        return;
                    }

                    if (ContactEntityForm.isNullOrEmpty(data)) {
                        return;
                    }
                    const startIndex = data.indexOf('{');
                    const endIndex = data.lastIndexOf('}');
                    if (startIndex < 0) {
                        return;
                    }

                    processResult.result = this.mapEntities(data.substring(startIndex, endIndex + 1));
                })
                .catch((error: Error) => {
                    processResult.error = error;
                })
                .finally(() => {
                    callback(processResult);
                });
        }

        static isNullOrEmpty(value: any): boolean {
            if (value === null || value === undefined) {
                return true;
            }

            if (typeof value === 'string' && value.trim() === '') {
                return true;
            }

            if (Array.isArray(value)) {
                return value.length === 0;
            }

            return false;
        }

        private getFetchQueryBody(entitySetName: string, fetchXml: string): string {
            const body: string[] = [];
            body.push('--batch_fetchquery');
            body.push(this.getFetchSet(entitySetName, fetchXml));
            body.push('--batch_fetchquery--');
            return body.join('\r\n');
        }

        private getFetchSet(entitySetName: string, fetchXml: string, batchId: string = null): string {
            const body: string[] = [];
            const fetchUrl = this.xrmApiUrl + entitySetName + '?fetchXml=' + encodeURIComponent(fetchXml);
            if (!ContactEntityForm.isNullOrEmpty(batchId)) {
                body.push('--batch_' + batchId);
            }

            body.push('Content-Type: application/http');
            body.push('Content-Transfer-Encoding:binary');
            body.push('');
            body.push('GET ' + fetchUrl + ' HTTP/1.1');
            body.push('Prefer: odata.include-annotations=*');
            body.push('Accept: application/json');
            body.push('');
            return body.join('\r\n');
        }

        private mapEntities(data: string): Entity[] {
            let resultEntities: Entity[];
            const entity = JSON.parse(data);
            resultEntities = entity.value;
            return resultEntities;
        }

        getHeaders(contentType: string) {
            const headers: any = {};
            headers.Accept = 'application/json';
            headers['OData-MaxVersion'] = '4.0';
            headers['OData-Version'] = '4.0';
            headers['Content-Type'] = contentType;
            headers.Prefer = 'odata.include-annotations="*"';
            return headers;
        }
    }

    export interface Entity {
        [propName: string]: any;
    }

    export class ProcessResult<T> {
        errorMessage: string | null;
        error: Error;
        result: T | null;

        constructor(result?: T | null) {
            this.result = result;
        }

        hasError(): boolean {
            return !ContactEntityForm.isNullOrEmpty(this.errorMessage);
        }
    }

    // eslint-disable-next-line @typescript-eslint/naming-convention
    export const ContactEntity = new ContactEntityForm();
}
