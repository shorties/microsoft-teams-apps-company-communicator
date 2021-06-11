import { ArrowRightIcon, TrashCanIcon } from '@fluentui/react-icons-northstar';
import { Button, Dropdown, Flex, Image, Label, List, Text, Loader } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";
import { TFunction } from "i18next";
import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import { RouteComponentProps } from 'react-router-dom';
import { createGroupAssociation, searchGroups, getGroupAssociations, deleteGroupAssociation } from "../../apis/messageListApi";
import { ImageUtil } from '../../utility/imageutility';
import './ManageGroups.scss';

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: {
        id: string,
    },
}

export interface IGroup {
    GroupId: string,
    GroupName: string,
    GroupEmail: string,
    ChannelId?: string,
}

export interface formState {
    loading: boolean,
    loader: boolean,
    channelId?: string, //id of the channel where the message was created
    channelName?: string,
    teamName?: string,
    userPrincipalName?: string,
    groups?: any[],
    groupAccess: boolean,
    noResultMessage: string,
    selectedGroups: dropdownItem[],
    selectedGroupsNum: number,
    allGroups: dropdownItem[],
    allGroupsNum: number,
    groupAlreadyIncluded: boolean,
}

export interface IManageGroupsProps extends RouteComponentProps, WithTranslation {
}

class ManageGroups extends React.Component<IManageGroupsProps, formState> {
    readonly localize: TFunction;
    targetingEnabled: boolean; // property to store value indicating if the targeting mode is enabled or not
    masterAdminUpns: string; // property to store value with the master admins

    constructor(props: IManageGroupsProps) {
        super(props);
        this.localize = this.props.t;
        this.targetingEnabled = false; // by default targeting is disabled
        this.masterAdminUpns = "";
        this.state = {
            loading: false,
            loader: true,
            channelId: "",
            channelName: "",
            teamName: "",
            userPrincipalName: "",
            groupAccess: false,
            noResultMessage: "",
            groupAlreadyIncluded: false,
            selectedGroups: [],
            selectedGroupsNum: 0,
            allGroups: [],
            allGroupsNum: 0,
        }
        this.escFunction = this.escFunction.bind(this);
    }

    public componentDidMount() {
        const setState = this.setState.bind(this);

        microsoftTeams.initialize();
        document.addEventListener("keydown", this.escFunction, false);

        microsoftTeams.getContext(context => {
            setState({
                channelId: context.channelId,
                channelName: context.channelName,
                teamName: context.teamName,
                userPrincipalName: context.userPrincipalName
            });

            //get all associated groups and set the allGroups and allGroupsNum state
            this.getAllGroupsAssociated();
        });
        
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public render(): JSX.Element 
    {
        return (
            <div>
                {(this.state.loader) &&
                    <div className="Loader">
                        <Loader />
                    </div>}
                { (!this.state.loader) &&
                    this.renderPage()}
            </div>
        );
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    private renderPage = () => {
        return (
            <div className="taskModule">
                <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                    <Flex className="nonScrollableContent">
                        <Flex.Item size="size.half">
                            <Flex column className="formContentContainer">
                                <div style={{ minHeight: 30 }} />
                                <div style={{ minHeight: 40 }}>
                                    <Label circular content={this.state.teamName} />
                                    <Label circular content={this.state.channelName} />
                                </div>
                                <div>
                                    <Flex gap="gap.small">
                                        <Dropdown
                                            search
                                            placeholder={this.localize("SendToGroupsPlaceHolder")}
                                            loadingMessage={this.localize("LoadingText")}
                                            onSearchQueryChange={this.onGroupSearchQueryChange}
                                            noResultsMessage={this.state.noResultMessage}
                                            loading={this.state.loading}
                                            items={this.getGroupItems()}
                                            onChange={this.onGroupsChange}
                                            value={this.state.selectedGroups}
                                            multiple
                                        />
                                        <Flex.Item><Button content="Add" icon={<ArrowRightIcon />} iconPosition="after" text onClick={this.onAddGroups} /></Flex.Item>
                                    </Flex>
                                </div>
                                <div className={this.state.groupAlreadyIncluded ? "ErrorMessage" : "hide"}>
                                    <div className="noteText">
                                        <Text error content={this.localize('GroupAlreadyIncluded')} />
                                    </div>
                                </div>
                            </Flex>
                        </Flex.Item>
                        <Flex.Item size="size.half">
                            <div className="scrollableContent">
                                <List items={this.state.allGroups} selectable />
                            </div>
                        </Flex.Item>
                    </Flex>
                </Flex>
                <Flex className="footerContainer" vAlign="end" hAlign="end">
                    <Flex className="buttonContainer" gap="gap.medium">
                        <Button content={this.localize('CloseText')} onClick={this.onClose} />
                    </Flex>
                </Flex>
            </div>
        );
    }

    private onClose() {
        microsoftTeams.tasks.submitTask();
    }

    private getGroupItems() {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        const dropdownItems: dropdownItem[] = [];
        return dropdownItems;
    }

    private makeDropdownItems = (items: any[] | undefined) => {
        const resultedTeams: dropdownItem[] = [];
        if (items) {
            items.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id,
                    },
                });
            });
        }
        return resultedTeams;
    }

    //executed when a group is added from the combo to the list
    private onAddGroups = () => {
        //for each one of the selected groups
        this.state.selectedGroups.forEach((element) => {
            //create a draft group based on IGroup interface
            var draftGroup: IGroup = {
                GroupId: element.key,
                GroupName: element.header,
                GroupEmail: element.content,
                ChannelId: this.state.channelId,
            }
            //If the group is not already on the list of associated groups, 
            //add the draftGroup to the database calling the webservice
            if (!this.state.allGroups.some(e => e.key === element.key)) {
                //add to the database
                this.saveGroup(draftGroup).then(() => {
                    //clears the combo box with selected groups
                    this.setState({
                        selectedGroups: [],
                        selectedGroupsNum: 0,
                    });

                    //refresh the list of associated groups
                    this.getAllGroupsAssociated();
                });
                //inputItems.push(draftGroup); //temporary, need to call the web service
            } else {
                this.setState({
                    groupAlreadyIncluded: true,
                });
            }
        });
    }

    //called to delete a group from the list
    private onDeleteGroup(id: number, key: string) {
        //removes from the list
        //this.state.allGroups.splice(id, 1);
        this.deleteGroup(key).then(() => {
            this.getAllGroupsAssociated();
        });
    }

    private deleteGroup = async (key: string) => {
        try {
            await deleteGroupAssociation(key);
        } catch (error) {
            return error;
        }
    }

    private saveGroup = async (draftGroup: IGroup) =>
    {
        try {
            await createGroupAssociation(draftGroup);
        } catch (error) {
            return error;
        }
    }

    private getAllGroupsAssociated = async () => {
        var resultListItems: any[] = [];

        try {
            //get inputGroups from database
            const response = await getGroupAssociations(this.state.channelId);
            const inputGroups = response.data;
            var x = 0;
            inputGroups.forEach((element) => {
                resultListItems.push({
                    id: x,
                    key: element.groupId,
                    header: element.groupName,
                    content: element.groupEmail,
                    endMedia: <Button circular size="small" onClick={this.onDeleteGroup.bind(this, x, element.groupId)} icon={<TrashCanIcon />} />,
                    media: <Image src={ImageUtil.makeInitialImage(element.groupName)} avatar />
                });
                x++;
            });

            this.setState({
                allGroups: resultListItems,
                allGroupsNum: resultListItems.length,
                loader: false,
            });
        } catch (error) {
            return error;
        }
    }

    private onGroupsChange = (event: any, itemsData: any) => {
        this.setState({
            selectedGroups: itemsData.value,
            selectedGroupsNum: itemsData.value.length,
            groups: [],
            groupAlreadyIncluded: false,
        })
    }

    private onGroupSearchQueryChange = async (event: any, itemsData: any) => {

        if (!itemsData.searchQuery) {
            this.setState({
                groups: [],
                noResultMessage: "",
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            this.setState({
                loading: false,
                noResultMessage: this.localize("NoMatchMessage"),
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            // handle event trigger on item select.
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string; }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            )
            if (result) {
                return;
            }

            this.setState({
                loading: true,
                noResultMessage: "",
            });

            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query);
                this.setState({
                    groups: response.data,
                    loading: false,
                    noResultMessage: this.localize("NoMatchMessage")
                });
            }
            catch (error) {
                return error;
            }
        }
    }
}

const manageGroupsWithTranslation = withTranslation()(ManageGroups);
export default manageGroupsWithTranslation;