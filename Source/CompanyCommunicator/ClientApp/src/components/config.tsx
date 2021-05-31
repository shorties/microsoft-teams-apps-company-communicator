// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../configVariables';
import { getAppSettings } from "../apis/messageListApi";
import { Loader, Checkbox } from '@fluentui/react-northstar';

export interface IConfigState {
    url: string;
    loading: boolean;
    channelId?: string;
    channelName?: string;
    teamName?: string;
}

class Configuration extends React.Component<{}, IConfigState> {
    targetingEnabled: boolean; // property to store value indicating if the targeting mode is enabled or not
    constructor(props: {}) {
        super(props);
        this.targetingEnabled = false; // by default targeting is disabled
        this.state = {
            url: getBaseUrl() + "/messages?locale={locale}",
            loading: true,
            channelId: "",
            channelName: "",
            teamName: ""

        }
    }

    public componentDidMount() {
        const setState = this.setState.bind(this);      
        microsoftTeams.initialize();
        
        microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
            microsoftTeams.settings.setSettings({
                entityId: "Company_Communicator_App",
                contentUrl: this.state.url,
                suggestedDisplayName: "Company Communicator",
            });
            saveEvent.notifySuccess();
        });

        // get the app settings and based on the targeting configuration and user id 
        // decides if the save is enabled or not
        this.getAppSettings().then(function () {
            setState({ loading: false });
            microsoftTeams.getContext(context => {
                setState({
                    channelId: context.channelId,
                    channelName: context.channelName,
                    teamName: context.teamName
                });
            });
            microsoftTeams.settings.setValidityState(true);
        });
    }

    public render(): JSX.Element {
        return (
            <div className="configContainer">
                {(this.state.loading) &&
                    <Loader label="Loading..." />}
                {(!this.state.loading) && this.renderTargetingMessage()}
            </div>
        );
    }

    // get the app configuration values and set targeting mode from app settings
    private getAppSettings = async () => {
        let response = await getAppSettings();
        if (response.data) {
            this.targetingEnabled = (response.data.targetingEnabled === 'true');
        }
    }

    // renders the message based on targeting configuration
    private renderTargetingMessage = () => {
        // check if targeting is enabled
        if (this.targetingEnabled) {
            // TODO
            // check if user is master admin (only master admins can install the authors app)
            return (
                <div>
                    <h3>You are configuring Company Communicator in target mode.</h3>
                    <p>
                        Please check below the team and channel where the Authors app is being installed:
                    </p>    
                    <p> 
                        {this.state.teamName} / {this.state.channelName} ({this.state.channelId})
                    </p>
                    <Checkbox
                        labelPosition="start"
                        label="Filter messages for this channel?"
                        checked={true}
                        toggle
                    />
                </div>
            )
        } else {
            return (
                <div>
                    <h3>Please click Save to get started.</h3>
                </div>
            )
        }
    }

}

export default Configuration;
