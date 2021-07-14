"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var react_icons_northstar_1 = require("@fluentui/react-icons-northstar");
var react_northstar_1 = require("@fluentui/react-northstar");
var microsoftTeams = require("@microsoft/teams-js");
var React = require("react");
var react_i18next_1 = require("react-i18next");
var messageListApi_1 = require("../../apis/messageListApi");
var imageutility_1 = require("../../utility/imageutility");
require("./ManageGroups.scss");
var react_image_file_resizer_1 = require("react-image-file-resizer");
var ManageGroups = /** @class */ (function (_super) {
    __extends(ManageGroups, _super);
    function ManageGroups(props) {
        var _this = _super.call(this, props) || this;
        //Function calling a click event on a hidden file input
        _this.handleUploadClick = function (event) {
            //reset the error message and the image link as the upload will reset them potentially
            _this.setState({
                errorImageUrlMessage: "",
                imageLink: ""
            });
            //fire the fileinput click event and run the handleimageselection function
            _this.fileInput.current.click();
        };
        _this.renderPage = function () {
            return (React.createElement("div", { className: "taskModule" },
                React.createElement(react_northstar_1.Flex, { column: true, className: "formContainer", vAlign: "stretch", gap: "gap.small", styles: { background: "white" } },
                    React.createElement(react_northstar_1.Flex, { className: "nonScrollableContent" },
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.half" },
                            React.createElement(react_northstar_1.Flex, { column: true, className: "formContentContainer" },
                                React.createElement("div", { style: { minHeight: 30 } }),
                                React.createElement("div", { style: { minHeight: 40 } },
                                    React.createElement(react_northstar_1.Label, { circular: true, content: _this.state.teamName }),
                                    React.createElement(react_northstar_1.Label, { circular: true, content: _this.state.channelName })),
                                React.createElement("div", null,
                                    React.createElement(react_northstar_1.Text, { content: _this.localize("CardImage") })),
                                React.createElement("div", { style: { minHeight: 100, maxHeight: 100, minWidth: 100, maxWidth: 100 } },
                                    React.createElement(react_northstar_1.Image, { fluid: true, src: _this.state.imageLink })),
                                React.createElement("div", { style: { minHeight: 40 } },
                                    React.createElement(react_northstar_1.Flex, { gap: "gap.smaller", vAlign: "end", className: "inputField" },
                                        React.createElement(react_northstar_1.Input, { value: _this.state.imageLink, placeholder: _this.localize("ImageURLPlaceHolder"), onChange: _this.onImageLinkChanged, error: !(_this.state.errorImageUrlMessage === ""), autoComplete: "off", fluid: true }),
                                        React.createElement("input", { type: "file", accept: "image/", style: { display: 'none' }, onChange: _this.handleImageSelection, ref: _this.fileInput }),
                                        React.createElement(react_northstar_1.Flex.Item, { push: true },
                                            React.createElement(react_northstar_1.Button, { circular: true, onClick: _this.handleUploadClick, size: "small", icon: React.createElement(react_icons_northstar_1.FilesUploadIcon, null), title: _this.localize("UploadImage") })))),
                                React.createElement("div", { style: { minHeight: 60 } },
                                    React.createElement(react_northstar_1.Input, { value: _this.state.channelTitle, onChange: _this.onChannelTitleChange, label: _this.localize("CardTitle"), fluid: true })),
                                React.createElement("div", null,
                                    React.createElement(react_northstar_1.Text, { content: _this.localize("TargetGroups") }),
                                    React.createElement(react_northstar_1.Flex, { gap: "gap.small" },
                                        React.createElement(react_northstar_1.Dropdown, { search: true, placeholder: _this.localize("SendToGroupsPlaceHolder"), loadingMessage: _this.localize("LoadingText"), onSearchQueryChange: _this.onGroupSearchQueryChange, noResultsMessage: _this.state.noResultMessage, loading: _this.state.loading, items: _this.getGroupItems(), onChange: _this.onGroupsChange, value: _this.state.selectedGroups, multiple: true }),
                                        React.createElement(react_northstar_1.Flex.Item, null,
                                            React.createElement(react_northstar_1.Button, { content: "Add", icon: React.createElement(react_icons_northstar_1.ArrowRightIcon, null), iconPosition: "after", text: true, onClick: _this.onAddGroups })))),
                                React.createElement("div", { className: _this.state.groupAlreadyIncluded ? "ErrorMessage" : "hide" },
                                    React.createElement("div", { className: "noteText" },
                                        React.createElement(react_northstar_1.Text, { error: true, content: _this.localize('GroupAlreadyIncluded') }))))),
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.half" },
                            React.createElement("div", null,
                                React.createElement(react_northstar_1.Text, { align: "center", content: _this.localize("TargetGroups") + ' for ' + _this.state.teamName + '/' + _this.state.channelName }),
                                React.createElement("div", { className: "scrollableContent" },
                                    React.createElement(react_northstar_1.List, { items: _this.state.allGroups, selectable: true })))))),
                React.createElement(react_northstar_1.Flex, { className: "footerContainer", vAlign: "end", hAlign: "end" },
                    React.createElement(react_northstar_1.Flex, { className: "buttonContainer", gap: "gap.medium" },
                        React.createElement(react_northstar_1.Button, { content: _this.localize('CloseText'), onClick: _this.onClose })))));
        };
        _this.onClose = function () {
            //collects values from state and build the draftChannel
            var draftChannel = {
                ChannelImage: _this.state.imageLink,
                ChannelId: _this.state.channelId,
                ChannelTitle: _this.state.channelTitle
            };
            //update the channel configuration and submit the task 
            _this.UpdateChannelConfig(draftChannel).then(function () {
                microsoftTeams.tasks.submitTask();
            });
        };
        //update or create a new channel configuration based on draftChannel
        _this.UpdateChannelConfig = function (draftChannel) { return __awaiter(_this, void 0, void 0, function () {
            var error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, messageListApi_1.updateChannelConfig(draftChannel)];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        return [2 /*return*/, error_1];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        //get the channel configuration 
        _this.GetChannelInfo = function (channelid) { return __awaiter(_this, void 0, void 0, function () {
            var response, draftChannel, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, messageListApi_1.getChannelConfig(channelid)];
                    case 1:
                        response = _a.sent();
                        draftChannel = response.data;
                        this.setState({
                            imageLink: draftChannel.channelImage,
                            channelTitle: draftChannel.channelTitle,
                        });
                        return [3 /*break*/, 3];
                    case 2:
                        error_2 = _a.sent();
                        return [2 /*return*/, error_2];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this.onChannelTitleChange = function (event) {
            _this.setState({
                channelTitle: event.target.value,
            });
        };
        _this.onImageLinkChanged = function (event) {
            var url = event.target.value.toLowerCase();
            if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
                _this.setState({
                    errorImageUrlMessage: _this.localize("ErrorURLMessage")
                });
            }
            else {
                _this.setState({
                    errorImageUrlMessage: ""
                });
            }
            _this.setState({
                imageLink: event.target.value,
            });
        };
        _this.makeDropdownItems = function (items) {
            var resultedTeams = [];
            if (items) {
                items.forEach(function (element) {
                    resultedTeams.push({
                        key: element.id,
                        header: element.name,
                        content: element.mail,
                        image: imageutility_1.ImageUtil.makeInitialImage(element.name),
                        team: {
                            id: element.id,
                        },
                    });
                });
            }
            return resultedTeams;
        };
        //executed when a group is added from the combo to the list
        _this.onAddGroups = function () {
            //for each one of the selected groups
            _this.state.selectedGroups.forEach(function (element) {
                //create a draft group based on IGroup interface
                var draftGroup = {
                    GroupId: element.key,
                    GroupName: element.header,
                    GroupEmail: element.content,
                    ChannelId: _this.state.channelId,
                };
                //If the group is not already on the list of associated groups, 
                //add the draftGroup to the database calling the webservice
                if (!_this.state.allGroups.some(function (e) { return e.key === element.key; })) {
                    //add to the database
                    _this.saveGroup(draftGroup).then(function () {
                        //clears the combo box with selected groups
                        _this.setState({
                            selectedGroups: [],
                            selectedGroupsNum: 0,
                        });
                        //refresh the list of associated groups
                        _this.getAllGroupsAssociated();
                    });
                    //inputItems.push(draftGroup); //temporary, need to call the web service
                }
                else {
                    _this.setState({
                        groupAlreadyIncluded: true,
                    });
                }
            });
        };
        _this.deleteGroup = function (key) { return __awaiter(_this, void 0, void 0, function () {
            var error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, messageListApi_1.deleteGroupAssociation(key)];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_3 = _a.sent();
                        return [2 /*return*/, error_3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this.saveGroup = function (draftGroup) { return __awaiter(_this, void 0, void 0, function () {
            var error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, messageListApi_1.createGroupAssociation(draftGroup)];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_4 = _a.sent();
                        return [2 /*return*/, error_4];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this.getAllGroupsAssociated = function () { return __awaiter(_this, void 0, void 0, function () {
            var resultListItems, response, inputGroups, x, error_5;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        resultListItems = [];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, messageListApi_1.getGroupAssociations(this.state.channelId)];
                    case 2:
                        response = _a.sent();
                        inputGroups = response.data;
                        x = 0;
                        inputGroups.forEach(function (element) {
                            resultListItems.push({
                                id: x,
                                key: element.groupId,
                                header: element.groupName,
                                content: element.groupEmail,
                                endMedia: React.createElement(react_northstar_1.Button, { circular: true, size: "small", onClick: _this.onDeleteGroup.bind(_this, x, element.groupId), icon: React.createElement(react_icons_northstar_1.TrashCanIcon, null) }),
                                media: React.createElement(react_northstar_1.Image, { src: imageutility_1.ImageUtil.makeInitialImage(element.groupName), avatar: true })
                            });
                            x++;
                        });
                        this.setState({
                            allGroups: resultListItems,
                            allGroupsNum: resultListItems.length,
                            loader: false,
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        error_5 = _a.sent();
                        return [2 /*return*/, error_5];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        _this.onGroupsChange = function (event, itemsData) {
            _this.setState({
                selectedGroups: itemsData.value,
                selectedGroupsNum: itemsData.value.length,
                groups: [],
                groupAlreadyIncluded: false,
            });
        };
        _this.onGroupSearchQueryChange = function (event, itemsData) { return __awaiter(_this, void 0, void 0, function () {
            var result, query, response, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!!itemsData.searchQuery) return [3 /*break*/, 1];
                        this.setState({
                            groups: [],
                            noResultMessage: "",
                        });
                        return [3 /*break*/, 6];
                    case 1:
                        if (!(itemsData.searchQuery && itemsData.searchQuery.length <= 2)) return [3 /*break*/, 2];
                        this.setState({
                            loading: false,
                            noResultMessage: this.localize("NoMatchMessage"),
                        });
                        return [3 /*break*/, 6];
                    case 2:
                        if (!(itemsData.searchQuery && itemsData.searchQuery.length > 2)) return [3 /*break*/, 6];
                        result = itemsData.items && itemsData.items.find(function (item) { return item.header.toLowerCase() === itemsData.searchQuery.toLowerCase(); });
                        if (result) {
                            return [2 /*return*/];
                        }
                        this.setState({
                            loading: true,
                            noResultMessage: "",
                        });
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 5, , 6]);
                        query = encodeURIComponent(itemsData.searchQuery);
                        return [4 /*yield*/, messageListApi_1.searchGroups(query)];
                    case 4:
                        response = _a.sent();
                        this.setState({
                            groups: response.data,
                            loading: false,
                            noResultMessage: this.localize("NoMatchMessage")
                        });
                        return [3 /*break*/, 6];
                    case 5:
                        error_6 = _a.sent();
                        return [2 /*return*/, error_6];
                    case 6: return [2 /*return*/];
                }
            });
        }); };
        _this.localize = _this.props.t;
        _this.targetingEnabled = false; // by default targeting is disabled
        _this.masterAdminUpns = "";
        _this.state = {
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
            imageLink: "",
            errorImageUrlMessage: "",
            channelTitle: "",
        };
        _this.escFunction = _this.escFunction.bind(_this);
        _this.fileInput = React.createRef();
        _this.handleImageSelection = _this.handleImageSelection.bind(_this);
        return _this;
    }
    ManageGroups.prototype.componentDidMount = function () {
        var _this = this;
        var setState = this.setState.bind(this);
        microsoftTeams.initialize();
        document.addEventListener("keydown", this.escFunction, false);
        microsoftTeams.getContext(function (context) {
            setState({
                channelId: context.channelId,
                channelName: context.channelName,
                teamName: context.teamName,
                userPrincipalName: context.userPrincipalName
            });
            //get all associated groups and set the allGroups and allGroupsNum state
            _this.getAllGroupsAssociated();
            //get the channel configuration from the database
            _this.GetChannelInfo(context.channelId);
        });
    };
    ManageGroups.prototype.componentWillUnmount = function () {
        document.removeEventListener("keydown", this.escFunction, false);
    };
    ManageGroups.prototype.render = function () {
        return (React.createElement("div", null,
            (this.state.loader) &&
                React.createElement("div", { className: "Loader" },
                    React.createElement(react_northstar_1.Loader, null)),
            (!this.state.loader) &&
                this.renderPage()));
    };
    ManageGroups.prototype.escFunction = function (event) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    };
    //function to handle the selection of the OS file upload box
    ManageGroups.prototype.handleImageSelection = function () {
        var _this = this;
        //get the first file selected
        var file = this.fileInput.current.files[0];
        if (file) { //if we have a file
            //resize the image to fit in the adaptivecard
            react_image_file_resizer_1.default.imageFileResizer(file, 100, 100, 'JPEG', 100, 0, function (uri) {
                if (uri.toString().length < 30720) {
                    //lets set the state with the image value
                    _this.setState({
                        imageLink: uri.toString()
                    });
                }
                else {
                    //images bigger than 32K cannot be saved, set the error message to be presented
                    _this.setState({
                        errorImageUrlMessage: _this.localize("ErrorImageTooBig")
                    });
                }
            }, 'base64'); //we need the image in base64
        }
    };
    ManageGroups.prototype.getGroupItems = function () {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        var dropdownItems = [];
        return dropdownItems;
    };
    //called to delete a group from the list
    ManageGroups.prototype.onDeleteGroup = function (id, key) {
        var _this = this;
        //removes from the list
        //this.state.allGroups.splice(id, 1);
        this.deleteGroup(key).then(function () {
            _this.getAllGroupsAssociated();
        });
    };
    return ManageGroups;
}(React.Component));
var manageGroupsWithTranslation = react_i18next_1.withTranslation()(ManageGroups);
exports.default = manageGroupsWithTranslation;
//# sourceMappingURL=ManageGroups.js.map