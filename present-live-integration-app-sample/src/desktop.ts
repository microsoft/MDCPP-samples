import { SpoFilePicker } from "./spoPicker";
const { v4: uuidv4 } = require('uuid');
import * as MicrosoftMeetings from "@microsoft/document-collaboration-sdk/";
import { getTokenImpl } from "./aadauth";


export function launch()
{
    const searchParams = new URLSearchParams( window.location.search );

    const sharingToken = searchParams.get( "ssToken" );
    const meetingId = searchParams.get( "meetingId" );
    if( sharingToken != null && meetingId != null )
    {
       loadAttendeeExperience( sharingToken, meetingId );
    }
    else
    {
        loadPresenterExperience();
    }
    
}

function loadPresenterExperience()
{
    MicrosoftMeetings.initializeMeeting( {
        hostName: "PPTLivePrototype",
        hostClientType: "Desktop"
    });

    const meetingId = uuidv4();
    const meetingInfo: MicrosoftMeetings.MeetingInfo = {
        id: meetingId
    };

    const userInfo: MicrosoftMeetings.UserInfo = {
        userId: 'prototype-unknownUser',
        participantId: uuidv4(),
        displayName: "User1"
    };

    const sessionInfo : MicrosoftMeetings.SessionInfo = {
        hostCorrelationId: uuidv4()
    }
    const fileSelector = new SpoFilePicker(
        (_token: string, spoUrl: string, itemId: string) => {
            const docInfo : MicrosoftMeetings.DocumentInfo = {
                documentIdentifier: {
                    siteUrl: spoUrl,  
                    sourceDoc: itemId 
                },
            };

            const container = window.document.createElement("div");
            container.style.height = "100%";
            window.document.body.appendChild(container);
        
            MicrosoftMeetings.startPresentation( {
                meetingInfo,
                userInfo,
                documentInfo: docInfo,
                sessionInfo,
                container,
                fetchAccessToken: async (resourceUrl: string, claim?: string[]) => {
                    let scopes : string[] = [];
                    claim?.forEach((val) => {
                        scopes.push(`${resourceUrl}/${val}`);
                    })
                    let token = await getTokenImpl( { scopes } )
                    return token; }
            })
            .then((response: MicrosoftMeetings.BootInfo) => {
                // Load attendee experience. This is the point where the 3rd party host needs to plumb the joinInfo on their end to initialize attendee sessions.
                console.log(`PPTLiveAPI.Present API finished response.isBootSuccess=${response.isBootSuccess}, response.joinInfo=${response.joinInfo}, meetingId=${meetingId}, response.errorInfo=${response.errorInfo}` );
                window.open(`http://localhost:8080/desktop.html?ssToken=${response.joinInfo}&meetingId=${meetingId}`);
             })
             .catch((e : any) => {
                 console.log(e);
             })
        }
    );
    fileSelector.launchPicker();
}

function loadAttendeeExperience( sharingToken: string, meetingId: string )
{
    MicrosoftMeetings.initializeMeeting( {
        hostName: "PPTLivePrototype",
        hostClientType: "Desktop"
    });

    const meetingInfo: MicrosoftMeetings.MeetingInfo = {
        id: meetingId
    };

    const userInfo: MicrosoftMeetings.UserInfo = {
        userId: 'prototype_unknownUser',
        participantId: uuidv4(),
        displayName: "User1"
    };

    const sessionInfo : MicrosoftMeetings.SessionInfo = {
        hostCorrelationId: uuidv4()
    }

    const container = window.document.createElement("div");
    container.style.height = "100%";
    window.document.body.appendChild(container);

    MicrosoftMeetings.joinPresentation( {
        meetingInfo,
        userInfo,
        sessionInfo,
        container,
        joinInfo: sharingToken
    })
    .then( (response: MicrosoftMeetings.BootInfo) => {
        console.log(`Attendee join success: isBootSuccess=${response.isBootSuccess}, meetingId=${meetingId}, errorInfo=${response.errorInfo}`);
    })
    .catch((err: any) => {
        console.log(`Attendee join fail err=${err}`);
    })
}