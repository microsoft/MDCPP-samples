import { combine, getToken } from "./aadauth";

export class SpoFilePicker
{ 
    private m_origin : string;
    private m_callback: (detailsToken : string, spoUrl : string, itemId : string) => void;

    private m_win : Window | null = null;
    private m_port : MessagePort | null = null;
    private m_iframe : HTMLIFrameElement | null = null;
    

    public constructor(callback: (detailsToken : string, spoUrl : string, itemId : string) => void)
    {
        this.m_origin = window.location.origin;
        this.m_callback = callback;
    }

    public async launchPicker() {
        // the options we pass to the picker page through the querystring
        const params = {
            sdk: "8.0",
            entry: {
                oneDrive: {
                    files: {},
                }
            },
            authentication: {},
            messaging: {
                origin: this.m_origin,
                channelId: "27"
            },
            typesAndSources: {
                mode: "files",
                filters: ['.pptx'],
                pivots: {
                    oneDrive: true,
                    recent: true,
                },
            },
            selection: {
                mode: 'single'
            },
            commands: {
                pick: {
                    label: 'Present'
                }
            }
        };

        const graphToken = await getToken({
            resource: "",
            type: "Default",
        });

        console.log("graph token: ", graphToken)

        const response = await fetch('https://graph.microsoft.com/v1.0/me?$select=mySite', {
        method: 'GET',
        headers: {
            "Authorization": `bearer ${graphToken}`
        }
        });
        const siteResponse = await response.json();
        console.log("siteResponse: ", siteResponse);
        let resource;
        try {
            resource = ( new URL(siteResponse.mySite) ).hostname;
        } catch (e) {
            console.error("Error extracting SPO hostname ", e);
            return;
        }
        const baseUrl = `https://${resource}/`;

        console.log("spo mysite: ", siteResponse)
        const authToken = await getToken({
            resource: baseUrl,
            type: "SharePoint",
        });

        this.m_iframe = window.document.createElement('iframe');
        this.m_iframe.height = "100%";
        this.m_iframe.width = "100%";
        window.document.body.append(this.m_iframe);
        this.m_win = this.m_iframe.contentWindow;


        const queryString = new URLSearchParams({
            filePicker: JSON.stringify(params),
        });

        const url = combine(baseUrl, `_layouts/15/FilePicker.aspx?${queryString}`);

        console.log(`Navigating to filePicker url: ${url}`);

        const form = this.m_win!.document.createElement("form");
        form.setAttribute("action", url);
        form.setAttribute("method", "POST");
        this.m_win!.document.body.append(form);

        const input = this.m_win!.document.createElement("input");
        input.setAttribute("type", "hidden")
        input.setAttribute("name", "access_token");
        input.setAttribute("value", authToken);
        form.appendChild(input);

        form.submit();

        window.addEventListener("message", (event) => {

            if (event.source && event.source === this.m_win!) {

                const message = event.data;

                if (message.type === "initialize" && message.channelId === params.messaging.channelId) {
                    this.m_port = event.ports[0];
                    this.m_port.addEventListener("message", (ev:MessageEvent) => { this.messageListener(ev ); });
                    this.m_port.start();
                    this.m_port.postMessage({
                        type: "activate",
                    });
                }
            }
        });
    }

    public async messageListener(message : MessageEvent) {
        switch (message.data.type) {

            case "notification":
                console.log(`notification: ${message.data}`);
                break;

            case "command":

                this.m_port!.postMessage({
                    type: "acknowledge",
                    id: message.data.id,
                });

                const command = message.data.data;

                switch (command.command) {

                    case "authenticate":

                        const token = await getToken(command);
                        if (typeof token !== "undefined" && token !== null) {
                            this.m_port!.postMessage({
                                type: "result",
                                id: message.data.id,
                                data: {
                                    result: "token",
                                    token,
                                }
                            });
                        } else {
                            console.error(`Could not get auth token for command: ${JSON.stringify(command)}`);
                        }

                        break;

                    case "close":
                        this.m_win!.close();
                        break;

                    case "pick":

                    console.log(`Picked: ${JSON.stringify(command)}`);

                    const ids = command.items[0].sharepointIds;
                    const spoUrl = ids.siteUrl;
                    const itemId = ids.listItemUniqueId;
                    const spoUri = new URL(spoUrl);
                    const detailsToken = await getToken(
                        {
                            resource: `https://${spoUri.hostname}/`,
                            type: "SharePoint"
                        }
                    );

                    this.m_iframe!.style.display = "none";
                    this.m_callback(detailsToken, spoUrl, itemId);

                    this.m_port!.postMessage({
                        type: "result",
                        id: message.data.id,
                        data: {
                            result: "success",
                        },
                    });

                    this.m_win!.close();
                    break;

                    default:

                        console.warn(`Unsupported command: ${JSON.stringify(command)}`, 2);
                        this.m_port!.postMessage({
                            result: "error",
                            error: {
                                code: "unsupportedCommand",
                                message: command.command
                            },
                            isExpected: true,
                        });
                        break;
                }

                break;
        }
    }
}