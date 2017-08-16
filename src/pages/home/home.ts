import { Component } from '@angular/core';
import { NavController } from 'ionic-angular';
import {
  ConfigurationApi, ExchangeService,UserSettingName, AutodiscoverService, Folder, FindItemsResults, SearchFilter, LogicalOperator, EmailMessageSchema,
  ItemView, FolderView, PropertySet, BasePropertySet, Item, ExchangeVersion, Uri, ImpersonatedUserId,
  Rule, FolderId, WellKnownFolderName, CreateRuleOperation, EwsLogging, ExchangeCredentials, WebCredentials,
  EmailMessage, MessageBody, OffsetBasePoint, Mailbox, ViewBase, ItemSchema, SortDirection, DateTime, Contact, ConnectingIdType } from "ews-javascript-api";
import { ntlmAuthXhrApi, cookieAuthXhrApi } from "ews-javascript-api-auth";

@Component({
  selector: 'page-home',
  templateUrl: 'home.html'
})

export class HomePage {

  constructor(public navCtrl: NavController) {
    ConfigurationApi.ConfigureXHR(new ntlmAuthXhrApi('emailid', 'password', true));
  }
  
  getOutlookData(){
    let exch = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
    exch.Credentials = new WebCredentials('emailid', 'password');
    exch.Url = new Uri("https://ews.domain/Ews/Exchange.asmx");
    exch.FindItems(WellKnownFolderName.Inbox, new ItemView(20)).then(
      response => {
        console.log(response);
      }).catch(function(error) {
        console.log(error);
      });
  }
}
