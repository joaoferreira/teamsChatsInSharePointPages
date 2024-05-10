/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-debugger */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prefer-const */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import Chat from "../Components/Chat/Chat";
import ChatNoPicture from "../Components/ChatNoPicture/ChatNoPicture";
import * as React from "react";
import * as ReactDOM from "react-dom";

import { graphfi, SPFx } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/photos";


import * as strings from 'TeamsChatEmbeddedApplicationCustomizerStrings';

import { app } from "@microsoft/teams-js";

export interface ITeamsChatEmbeddedApplicationCustomizerProperties {}



/** A Custom Action which can be run during execution of a Client Side Application */
export default class TeamsChatEmbeddedApplicationCustomizer
  extends BaseApplicationCustomizer<ITeamsChatEmbeddedApplicationCustomizerProperties> {
    private _bottomPlaceholder: PlaceholderContent | undefined;

    @override
    public async onInit(): Promise<void> { 
           
      try {
        //Detect if the SharePoint page is running inside Microsoft Teams 
        //If in Microsoft Teams end the excution
        await app.initialize();
        const context = await app.getContext();
        if(context){
          return;
        }
      } catch (exp) {

        this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
        let profilePictureUrl:string;
        //Get User Profile from Microsoft Graph to ensure the most updated profile picture 
        //If permission is not granted by the administrator fallback to the classic SharePoint profile picture
        try{
          const graph = graphfi().using(SPFx(this.context));
          const photoValue = await graph.me.photo.getBlob();      
          const url = window.URL || window.webkitURL;
          profilePictureUrl = url.createObjectURL(photoValue);
          //Render Chat component with the user profile picture
          const chat = React.createElement(Chat, { label: strings.Label, userPhoto: profilePictureUrl });
          ReactDOM.render(chat, this._bottomPlaceholder.domElement);  
        }catch(exPhoto){   
          //Render Chat component without the user profile picture
          const chatNoPicture = React.createElement(ChatNoPicture, { label: strings.Label });
          ReactDOM.render(chatNoPicture, this._bottomPlaceholder.domElement);  
        }
        
      }   

      return Promise.resolve();
    }


  }



