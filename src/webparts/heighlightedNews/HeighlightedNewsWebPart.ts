import { Version } from '@microsoft/sp-core-library';
require('sp-init');  
require('microsoft-ajax');  
require('sp-runtime');  
require('sharepoint'); 
require('sp-strings'); 
//require('sp-date-time'); 

import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { each, escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http'; 

import { PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy, IColumnReturnProperty } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import { PropertyFieldSearch } from '@pnp/spfx-property-controls/lib/PropertyFieldSearch';

import styles from './HeighlightedNewsWebPart.module.scss';
import * as strings from 'HeighlightedNewsWebPartStrings';

//import '~office-ui-fabric-react/dist/sass/_References.scss';
import { getIconClassName } from '@uifabric/styling';
import * as moment from 'moment';

export interface IHeighlightedNewsWebPartProps {
  newsSiteUrl: string;
  sourceSite: string;
  sourceLib:string;
  sourcePage:string;
  filterValue: string;
  filterColumn: string;
  list: string;
  searchValue: string;
  selectedValue: string;
  selectedPage: string;
  filterCondition: string;
  operatorType: string;
  showDate: boolean;
  showAuthor: boolean;
  isAdvSettingEnabled: boolean;
  requireTitleTruncate: boolean;
  requireDescTruncate: boolean;
}

export interface spListItems{  
  value: spListItem[];
}  

export interface spListItem{  
  FileLeafRef: string;  
  FileName: string;  
  AbsoluteUrl:string;
  UniqueId:string;
  File:{
    ServerRelativeUrl:string;
  };
  // Created: string;  
  // Author: {  
  //   Title: string;  
  // };

}  
  
export interface spList{  
Title:string;  
id: string; 
SiteId;
}

export interface spLists{  
  value: spList[];  
}

export interface propertyField{  
  Title:string;  
  StaticName: string;
  }
  
  export interface propertyFields{  
    value: propertyField[];  
  }

  export interface newsItem
  {  
    Title:string;  
    Created: string;  
    Author: {  
      Title: string;  
    };
  }
    
  export interface newsItems{  
    value: newsItem[];  
  }

export default class HeighlightedNewsWebPart extends BaseClientSideWebPart<IHeighlightedNewsWebPartProps> {
  private libOrSiteDropDownOptions: IPropertyPaneDropdownOption[] =[];  
  private pageDropDownOptions: IPropertyPaneDropdownOption[] = [];
  private filterDropDownOptions: IPropertyPaneChoiceGroupOption[] = [];
  private selectedFields: IPropertyPaneChoiceGroupOption[] = [];
  private selectedPages: IPropertyPaneChoiceGroupOption[] = [];
  private mainNews: any[] = [];
  private subNews: any[] = [];
  private propPaneFields: any [] = [];
  private propFilterButtons: any[] = [];
  private authorImg: string = "";
  private htmlScript: string = `
                                var itms = document.getElementsByClassName("dt-time");
                                for (var i = 0; i < itms.length; i++) {
                                  //console.log(itms[i].innerText);
                                  var hrs = new Date(itms[i].innerText).format("H:mm");
                                  var txt = SP.DateTimeUtil.SPRelativeDateTime.getRelativeDateTimeString(new Date(itms[i].innerText), true, SP.DateTimeUtil.SPCalendarType.none,false);
                                  var sArr = txt.split(' at ');
                                  if(sArr.length >1)
                                  {
                                    txt = sArr[0] + " at " + hrs;
                                  }
                                  itms[i].innerText = txt;
                                }
                                `;
  private enableFilterTextbox: boolean = true;
  private requireToLoadNews: boolean = true;
  private requireToRenderHTML: boolean = true;
  private isDateTimeJsRefAdded: boolean = false;
  private countArrIncrease: number = 2;// Search result column starts in insida from 2 index/ docs node 0
  private layOutType: number = 0;// 1-One Column 2-Two Column 3-Three Column
  private sector: string = "";

  public render(): void {
    let mainNewsContent: string = "";
    let subNewsContent: string = "";
    let newsImageAPIUrl = "";
    let siteUrl = "";
    moment.locale('nb');

    this.layOutType = this.context.domElement.offsetWidth <= 440 ? 3 :
    this.context.domElement.offsetWidth > 441 && 
    this.context.domElement.offsetWidth <= 500 ? 2 :
    this.context.domElement.offsetWidth > 1200 && 
    this.context.domElement.offsetWidth <= 2000 ? 1 : 0;
    console.log(this.layOutType);
    const cssClasses = styles.index + ' ' +  ' ' + (this.layOutType == 3  ? styles.xsConta : '') +  ' ' + (this.layOutType ==2 ? styles.smConta : '')  +  ' ' + (this.layOutType ==1 ? styles.xlConta : '');

    if (!this.isDateTimeJsRefAdded) {
      this.AddJSFileToHeader(this.context.pageContext.web.absoluteUrl + "/_layouts/15/SP.dateTimeUtil.js");
      this.isDateTimeJsRefAdded = true;
    }

    if (this.sector == "") 
    {
      //Get current user sector property
      this.GetJsonData(this.context.pageContext.web.absoluteUrl + 
                      "/_api/SP.UserProfiles.PeopleManager/GetMyProperties").then((response)=>{ 
        var rslt = response.value !== undefined ? response.value:response;
       
        if (rslt !== undefined && rslt.UserProfileProperties.length > 0) 
        {
          rslt.UserProfileProperties.forEach(element => {
           if (element.Key === "ORG-Sektor") {
             this.sector = element.Value;
           }
          });
        }
      });  
    }

    siteUrl = this.properties.sourceSite === "site" ? 
              this.context.pageContext.web.absoluteUrl :
              this.properties.sourceSite === "hub" ? 
              this.properties.sourceLib :
              this.properties.sourceSite === "news" ? 
              this.properties.newsSiteUrl:"";
    
    siteUrl = siteUrl === undefined ? "": 
              siteUrl.endsWith("/") ? siteUrl.slice(0,-1) : 
              siteUrl;

    newsImageAPIUrl = siteUrl + "/_layouts/15/getpreview.ashx?path=";

  // newsImageAPIUrl = siteUrl + "/_api/v2.0/sites/78dad885-9f2e-4102-9444-6352d72e5a59/lists/89a1b037-cf36-4c30-9de8-733a53e85923/items/";
   //b7fdc08a-42eb-40ec-a06d-c62c7744b9d2/driveItem/thumbnails/0/c300x400/content";
    
    if (this.requireToLoadNews) 
    {
      // Called for fresh data load, so fetch data
      this.PopulateNews();

      //Reset to false
      this.requireToLoadNews = false;
    }

    if (this.requireToRenderHTML) 
    {
      //Reset to false
      this.requireToRenderHTML = false;

      if(this.mainNews.length == 1)
      {
        let listresturlAA = this.properties.newsSiteUrl;
        let domainUrlAA = this.GetDominUrl();
        let fileNameAA= this.properties.sourcePage;

        let itemGuid = this.GetPageGuid(this.properties.sourceSite,listresturlAA,domainUrlAA,fileNameAA);
       // let mainImgUrl: string = newsImageAPIUrl + siteUrl + "/SitePages/" + this.properties.sourcePage + "&force=1&resolution=2";

       let mainImgUrl: string = listresturlAA + "/_api/v2.0/sites/b3d4d005-f37a-484d-9b02-190f8c50b0e4/lists/0e36dd35-b832-44f4-ac85-15f689f71723/items/{GUID}/driveItem/thumbnails/0/c300x400/content";

       console.log("MAin section Thumbnail: "+ mainImgUrl);
        
        if (this.mainNews[0].AuthorEmail !== "") 
        {
          this.authorImg = siteUrl + "/_layouts/15/userphoto.aspx?size=S&username=" + this.mainNews[0].AuthorEmail;
        }
       
        if (this.properties.sourceSite === "hub") 
        {
          mainImgUrl = this.mainNews[0].SiteUrl + "/_layouts/15/getpreview.ashx?path="+ this.mainNews[0].ViewUrl + "&force=1&resolution=2";  
          this.authorImg = this.mainNews[0].SiteUrl + "/_layouts/15/userphoto.aspx?size=S&username=" + this.mainNews[0].AuthorEmail;
        }

        var mainTitle = "", mainDesc = "";
        mainTitle = this.mainNews[0].Title !==null ? this.mainNews[0].Title:"" ;
        mainDesc = this.mainNews[0].Description !==null ? this.mainNews[0].Description: "";

        if (this.properties.requireTitleTruncate) 
        {
          if (mainTitle !== null && mainTitle !== "" && mainTitle.length > 70) 
          {
            mainTitle = this.mainNews[0].Title.substring(0, 70) + "...";
          }
        }

        if (this.properties.requireDescTruncate) 
        {
          if (mainDesc !== null && mainDesc !== undefined && mainDesc !== "" && mainDesc.length > 70) 
          {
            mainDesc= this.mainNews[0].Description.substring(0, 70) + "...";
          }
        }
                          //<div class="${styles.watermark}" ><i class="${getIconClassName('POI')}" ></i> ${this.mainNews[0].Sektor}</div>
        mainNewsContent = `<div class="${styles.imgBigcontainer}">
                             <a  href=" ${this.mainNews[0].ViewUrl}">
                              <img src="${mainImgUrl}" class="${styles.rsImage} ${styles.mBot10} ${styles.mt2}  " />
                              
                             </a>
                            </div>
                            <div class="${styles.contentContainer} ${styles.mainSecBackground}" style="margin-top:18px">
                              <div>
                                <a href=" ${this.mainNews[0].ViewUrl}" class="${styles.linkq}"> 
                                    ${mainTitle}
                                </a>
                              </div>
                              <div class="${styles.hd}">
                                ${mainDesc}
                              </div>
                              <div class="${styles.textMuted}">
                                
                              </div>
                              <div class="${styles.mTop10}">
                                <div class="${styles.media}">
                                  <img class="${styles.mr3} ${styles.prof} author" src="${this.authorImg}" />
                                    <div class="${styles["media-body"]}" style="padding-top:8px">
                                      <span class="${styles.ft12} author">
                                        <b>${this.mainNews[0].Author}</b>
                                      </span>
                                      <span class="${styles.textMuted} dt-time">
                                       ${moment(this.mainNews[0].Created,"YYYY-MM-DDTHH:mm:ss").fromNow()}
                                      </span>
                                      <a  href="https://vtfk.sharepoint.com/sites/innsida-siste-nytt/sitepages/news-hub.aspx" class="${styles.customButton}">Flere nyhetssaker</a>
                                    </div>
                              </div>
                              </div>
                            </div>`;
      }
      else
      {
        mainNewsContent = "Please set news page from property pane.";
      }
  
      if(this.subNews.length > 0)
      {
        let totSubNewsCount = 0;
  
        for (let subNewsCount = 0; subNewsCount < this.subNews.length; subNewsCount++) 
        {
          const currentSubNews = this.subNews[subNewsCount];
          let subNewsUrl = newsImageAPIUrl + siteUrl + "/SitePages/"
          var subTitle = "", subDesc = "";
          subTitle = currentSubNews.Title !== null ? currentSubNews.Title : "";
          subDesc = currentSubNews.Description !== null ? currentSubNews.Description : "";
          
          if ((this.mainNews[0] === undefined || 
               this.mainNews[0].Name !== currentSubNews.Name) 
               && totSubNewsCount < 3) 
          {
            if (this.properties.sourceSite === "hub") 
            {
              subNewsUrl = currentSubNews.SiteUrl + "/_layouts/15/getpreview.ashx?path="+ currentSubNews.ViewUrl + "&force=1&resolution=2";  
            }
            else
            {
             // subNewsUrl = subNewsUrl + currentSubNews.Name + ".aspx";

             subNewsUrl = siteUrl + "/_api/v2.0/sites/b3d4d005-f37a-484d-9b02-190f8c50b0e4/lists/0e36dd35-b832-44f4-ac85-15f689f71723/items/"+ currentSubNews.UniqueID +"/driveItem/thumbnails/0/c300x400/content";
            }

            if (this.properties.requireTitleTruncate) 
            {
              if (subTitle !== null && subTitle !== "" && subTitle.length > 70) 
              {
                subTitle = currentSubNews.Title.substring(0, 70) + "...";
              }
            }

            if (this.properties.requireDescTruncate) 
            {
              if (subDesc !== null && subDesc !== undefined && subDesc !== "" && subDesc.length > 70) 
              {
                subDesc = currentSubNews.Description.substring(0, 70) + "...";
              }
            }

            subNewsContent += `<div class="${styles.media} ${styles.mb2} ${styles.ctmed}">
                                  <div class="${styles.imgSMcontainer}">
                                  <a href="${currentSubNews.ViewUrl}"> 
                                    <img class="${styles.mr3}" src="${subNewsUrl}" />
                                    <div class="${styles.watermark}" ><i class="${getIconClassName('POI')}" ></i> ${currentSubNews.Sektor}</div>
                                  </a>
                                  </div>
                                    <div class="${styles["media-body"]}">
                                      <a href="${currentSubNews.ViewUrl}" class="${styles.linkq}"> 
                                        ${subTitle}
                                      </a>
                                      <div class="${styles.spHeader} ${styles.mb1}">
                                        ${subDesc}
                                      </div>
                                      <div class="${styles.textMuted} ${styles.mb1}">
                                        
                                      </div>
                                      <div class="${styles.ft12}">
                                        <span class="author">
                                          <b>
                                            ${currentSubNews.Author}
                                          </b>
                                        </span> 
                                        <span class="dt-time"> 
                                          ${moment(currentSubNews.Created,"YYYY-MM-DDTHH:mm:ss").fromNow()}
                                        </span>
                                      </div>
                                    </div>
                                </div>`;
            totSubNewsCount ++;
          }
        }
      }
      else
      {
        subNewsContent = "No news available";
      }
      this.GetPageGuid(this.properties.sourceSite,this.properties.newsSiteUrl,this.GetDominUrl(),this.properties.sourcePage).then((response)=>{ 
        var rslt = response.value !== undefined ? response.value:response;
        console.log("Result for GUID" );
        console.log(rslt);
        var abc="";

        if (rslt !== undefined && rslt.length > 0) 
          {
            abc=rslt[0].UniqueId;
            var mainNewsCntnt = mainNewsContent;
            mainNewsCntnt = mainNewsCntnt.replace("{GUID}",abc);
            this.domElement.innerHTML = `<div>
                                    <div class="${cssClasses}">
                                      <div class="">
                                        <div class="${ styles.row}">
                                          <div class="${ styles.col} ${styles.mBot10}">
                                            ${mainNewsCntnt}
                                          </div>
                                          <div class="${ styles.col}  ${styles.mt2}">
                                            ${subNewsContent}
                                          </div>
                                        </div>
                                      </div>
                                      </div>
                                    </div>
                                  `;
            
            // JS Script executed to set User friendly date value using sharepoint SP.DateTimeUtility method
        let head: any = document.getElementsByTagName("head")[0] || document.documentElement,
        script = document.createElement("script");
        script.type = "text/javascript";
  
        try 
        {
          var authorShowHide = "";
          var dateShowHide = "";
          if (this.properties.showAuthor === undefined || this.properties.showAuthor) 
          {
            authorShowHide = `
                                [].forEach.call(document.querySelectorAll('.author'), function (el) {
                                  el.style.display = 'inline';
                                });
                               `;
          }
          else
          {
            authorShowHide += `
                                [].forEach.call(document.querySelectorAll('.author'), function (el) {
                                  el.style.display = 'none';
                                });
                               `;
          }

          if (this.properties.showDate === undefined || this.properties.showDate) 
          {
            dateShowHide += `
                                [].forEach.call(document.querySelectorAll('.dt-time'), function (el) {
                                  el.style.display = 'inline';
                                });
                               `;
          }
          else
          {
            dateShowHide += `
                                [].forEach.call(document.querySelectorAll('.dt-time'), function (el) {
                                  el.style.display = 'none';
                                });
                               `;
          }
          
         // script.appendChild(document.createTextNode(this.htmlScript));
          script.appendChild(document.createTextNode(authorShowHide));
          script.appendChild(document.createTextNode(dateShowHide));
        } 
        catch (e) 
        {
          // script.text = this.htmlScript + authorShowHide + dateShowHide;
          script.text = authorShowHide + dateShowHide;
        }
        
        head.insertBefore(script, head.firstChild);
        head.removeChild(script);
           }
          }); 
    }
  }

  protected onPropertyPaneConfigurationStart(): void 
  {  
    if (this.properties.isAdvSettingEnabled === undefined || !this.properties.isAdvSettingEnabled) 
    {
      this.SetPropertyPaneFields(false);
      this.propFilterButtons = [];
    }
    else
    {
      this.SetPropertyPaneFields(true);
      //Render buttons
      this.GetButtonsForProperty(false);
    }

    this.context.propertyPane.refresh();

    // loads library/site names based on seleted source(site/hub)
    this.PopulateSiteOrHub();
    this.GetFilesFromLibOrHub();
    
    // //Populate properties for filter
    this.PopulateFilterProperties();
    //this.PopulateNews();
    this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
  }

  protected   onPropertyPaneFieldChanged ( propertyPath :   string , oldValue :   any , newValue :   any ) :   void  {
    
    if  (propertyPath === "sourceSite" || 
          propertyPath === "newsSiteUrl") 
    {
        this.PopulateSiteOrHub();
        this.ClearPropertyFields();
       
        this.requireToLoadNews = true;
        this.requireToRenderHTML = true;
    }
    else if(propertyPath === "isAdvSettingEnabled")
    {
      if (this.properties.isAdvSettingEnabled) 
      {
        this.SetPropertyPaneFields(true);

        //Render buttons
        this.GetButtonsForProperty(false);
      }
      else
      {
        this.propFilterButtons = [];

        this.SetPropertyPaneFields(false);
      }
    }
    else if(propertyPath === "sourceLib")
    {
        this.GetFilesFromLibOrHub();
      
        this.requireToLoadNews = true;
    }
    else if(propertyPath === "showDate")
    {
      this.requireToRenderHTML = true;
    }
    else if(propertyPath === "showAuthor")
    {
      this.requireToRenderHTML = true;
    }
    else if(propertyPath === "sourcePage")
    {
      this.selectedPages = [];

      if(this.properties.sourcePage.length > 3)
      {
        this.pageDropDownOptions.forEach(item => {
          
          if(item.text.toLocaleLowerCase().indexOf(this.properties.sourcePage.toLocaleLowerCase()) !== -1)
          {
            this.selectedPages.push(item);
          }
        });
      }

      this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
    }
    else if(propertyPath === "searchValue")
    {
      this.selectedFields = [];

      if(this.properties.searchValue.length > 3)
      {
        this.filterDropDownOptions.forEach(item => {
          
          if(item.text.toLocaleLowerCase().indexOf(this.properties.searchValue.toLocaleLowerCase()) !== -1)
          {
            this.selectedFields.push({key:item.key, text:item.text, checked: false});
          }
        });
      }

      if (this.selectedFields.length > 0) 
      {
        this.selectedFields.unshift({key:"1", text:"Select", checked: true});
      }

      this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
    }
    else if(propertyPath === "selectedValue")
    {
      this.properties.searchValue = this.properties.selectedValue;
      this.properties.selectedValue = "";
      this.selectedFields = [];

      this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
    }
    else if(propertyPath === "selectedPage")
    {
      this.properties.sourcePage = this.properties.selectedPage;
      
      this.requireToLoadNews = true;
      this.selectedPages = [];

      this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
    }
    else if  (propertyPath === "requireDescTruncate" || 
          propertyPath === "requireTitleTruncate") 
    {
      this.requireToRenderHTML = true;
    }

    this.context.propertyPane.refresh();
 }

 private PopulateSiteOrHub(): void
 {
    this.libOrSiteDropDownOptions = [];
    
    if(this.properties.sourceSite === "site" ||
       this.properties.sourceSite === "news"||
       (this.properties.sourceSite === undefined && 
        this.properties.newsSiteUrl !== undefined &&
        this.properties.newsSiteUrl.length > 20))
    {
      //this.GetAllLibraries();
      this.libOrSiteDropDownOptions.push({key:"Site Pages",text:"Site Pages"});
      this.properties.sourceLib = "Site Pages";
      this.GetFilesFromLibOrHub();
    }
    else if(this.properties.sourceSite === "hub")
    {
      this.GetHubSites();
    }
 }

  private GetAllLibraries():void{  
    // REST API to pull the library names  
    let listresturl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$select=Id,Title&filter=BaseTemplate eq 101";  
    
    this.LoadLibraryOrSiteDropDownData(listresturl).then((response)=>{  
      // Polulate property pane Libray/Site dropdown
      this.PopulateLibrayOrSiteDropDownValues(response.value);
    });
  } 

  private GetHubSites():void{  
    // REST API to pull the list names  
    let listresturl: string = this.context.pageContext.web.absoluteUrl + "/_api/HubSites?select=SiteId,Title";  
    
    console.log("Hub Sites: " + listresturl);

    this.LoadLibraryOrSiteDropDownData(listresturl).then((response)=>{  
      // Polulate property pane Libray/Site dropdown
      this.PopulateLibrayOrSiteDropDownValues(response.value);
    });
  } 

  private PopulateFilterProperties():void{  
    let siteUrl = this.properties.sourceSite === "site" ? 
                  this.context.pageContext.web.absoluteUrl :
                  this.properties.sourceSite === "hub" ? 
                  this.context.pageContext.web.absoluteUrl :
                  this.properties.sourceSite === "news" ? 
                  this.properties.newsSiteUrl:"";

    // REST API to pull the property columns  
    let listresturl: string = siteUrl + "/_api/web/lists/getByTitle('Site Pages')/fields?$select=Title,StaticName&$filter=Hidden eq false and FieldTypeKind gt 0";  

    if (siteUrl !== "") 
    {
      console.log("Fields: " + listresturl);
      this.LoadPropertyDropDownData(listresturl).then((response)=>{  
        // Polulate property pane Libray/Site dropdown
        this.PopulatePropertyDropDownValues(response.value);
      });
    }
  }

  private PopulateNews():void{ 
    if(this.properties.sourceSite !== undefined || 
      (this.properties.newsSiteUrl !== "" && 
      this.properties.newsSiteUrl !== undefined &&
      this.properties.newsSiteUrl.length > 20))
    {
        // REST API to pull the news
        let listresturl: string = "";
        let listresturlMainNews: string = "";
        let listresturlSubNews: string = "";

        // var fileUrl = this.properties.sourcePage !== undefined ? this.properties.sourcePage.split('/'): [""];
        // var fileName = fileUrl[fileUrl.length -1].split('.aspx')[0];
        var fileName = this.properties.sourcePage;
        var domainUrl = this.GetDominUrl();

        if(this.properties.sourceSite === "site")
        {
          listresturl = this.context.pageContext.web.absoluteUrl;
        }
        else if(this.properties.sourceSite === "hub")
        {
          listresturl = this.properties.sourceLib;
        }
        else if(this.properties.sourceSite === "news" || 
              (this.properties.sourceSite === undefined &&
               this.properties.newsSiteUrl !== undefined &&
               this.properties.newsSiteUrl.length > 20))
        {
          listresturl = this.properties.newsSiteUrl;
        }

        if(listresturl !== undefined && listresturl !== "")
        {
          listresturlMainNews = this.GetNewsUrl("main",this.properties.sourceSite,listresturl,domainUrl,fileName);
          listresturlSubNews = this.GetNewsUrl("sub",this.properties.sourceSite,listresturl,domainUrl,fileName);

          console.log("Main News: " + listresturlMainNews);

          if(fileName !== "")
          {
            //Get and fill array with main news
            this.GetJsonData(listresturlMainNews).then((response)=>{ 
              var rslt = response.value !== undefined ? response.value:response;

              // Polulate main news
              this.PopulateNewsValues(rslt, true);
            });
          }
          else
          {
            this.mainNews = [];
          }

          if (this.properties.filterCondition !== undefined) 
          {
            var filter = this.properties.filterCondition;

            if (filter.length > 0 && filter.charAt(0) == "'") 
            {
              listresturlSubNews = listresturlSubNews;

              console.log("Top News: " + listresturlSubNews);
          
              //Get and fill array with sub news
              this.GetJsonData(listresturlSubNews).then((response)=>{ 
                var rslt = response.value !== undefined ? response.value:response;
                // Polulate sub news
                this.PopulateNewsValues(rslt, false);
              });
            }
            else
            {
              if (filter.trim().length > 0) 
              {
                listresturlSubNews += "&$filter=" + filter; 
              }

              listresturlSubNews = listresturlSubNews.replace(/'true'/ig,"true").replace(/'false'/ig,"false");

              console.log("Top News: " + listresturlSubNews);
          
              //Get and fill array with sub news
              this.GetJsonData(listresturlSubNews).then((response)=>{ 
                var rslt = response.value !== undefined ? response.value:response;
                // Polulate sub news
                this.PopulateNewsValues(rslt, false);
              });
            }
          }
        }
        else
        {
          this.subNews = [];
          this.mainNews = [];
        }
    }
  }

  private PopulatePropertyDropDownValues(lists: propertyField[]): void{  
    
    this.filterDropDownOptions = [];

    lists.forEach((list:propertyField)=>{  
      // Loads the drop down values  
      this.filterDropDownOptions.push({key:list.StaticName,text:list.Title, checked:false}); 
    });

    this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
  }
  
  private PopulateLibrayOrSiteDropDownValues(lists: spList[]): void{  
    
    this.libOrSiteDropDownOptions = [];

    if(this.properties.sourceSite === "site" || 
       this.properties.sourceSite === "news")
      {
        lists.forEach((list:spList)=>{  
          // Loads the drop down values  
          this.libOrSiteDropDownOptions.push({key:list.Title,text:list.Title}); 
        });
      }
      else if(this.properties.sourceSite === "hub")
      {
        lists.forEach((list:spList)=>{  
          // Loads the drop down values  
          this.libOrSiteDropDownOptions.push({key:list.SiteId,text:list.Title}); 
        });
      }

    this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
  }

  private PopulateNewsValues(lists: any, isMainNews: boolean): void{  
    let items = [];
    let isHubItems:boolean = false;

    if (lists.PrimaryQueryResult == undefined) 
    {
      items = lists;
    }
    else
    {
      items = lists.PrimaryQueryResult.RelevantResults.Table.Rows;
      isHubItems = true;
    }
    
    if(isMainNews)
    {
      this.mainNews = [];
      
      if (items !== undefined && items.length>0) {
        items.forEach((list:any)=>{  
          // Loads the main news details  

          if (!isHubItems) {
            this.mainNews.push(
              {
                  Title:list.Title === undefined || list.Title === null ? "":list.Title, 
                  Created: list.ListItemAllFields === undefined ? "" : list.ListItemAllFields.Modified,
                  Author: list.Author === undefined ? "" : list.Author.Title,
                  AuthorEmail: list.Author === undefined ? "dummy@noimage.dummy" : list.Author.Email,
                  Name: list.Name === undefined || list.Name === null ? "" : list.Name.replace(".aspx", ""),
                  AuthorID: list.ListItemAllFields === undefined ? "0" : list.ListItemAllFields.AuthorId,
                  ViewUrl: list.ServerRelativeUrl === undefined ? 
                            list.AbsoluteUrl === undefined ? "#" : 
                            list.AbsoluteUrl : 
                            this.GetDominUrl() + list.ServerRelativeUrl,
                  Description: list.ListItemAllFields.Description === undefined || list.ListItemAllFields.Description === null ? "":list.ListItemAllFields.Description,
                  Sektor: list.ListItemAllFields.Sektor0 === undefined || list.ListItemAllFields.Sektor0 === null ? "":list.ListItemAllFields.Sektor0.Label
              }
            );  
          }
          else
          {
            var arr = list.Cells;

            this.mainNews.push(
              {
                  Title: arr[0 + this.countArrIncrease].Value, 
                  Created: arr[1 + this.countArrIncrease].Value,
                  Author: arr[3 + this.countArrIncrease].Value !== null ? arr[3 + this.countArrIncrease].Value.split('|')[1]:"-",
                  AuthorEmail: arr[3 + this.countArrIncrease].Value !== null ? arr[3 + this.countArrIncrease].Value.split('|')[0]:"noemail",
                  ViewUrl: arr[4 + this.countArrIncrease].Value,
                  SiteUrl: arr[5 + this.countArrIncrease].Value,
                  Name: this.GetPageNameFromUrl(arr[4 + this.countArrIncrease].Value),
                  Description: arr[6 + this.countArrIncrease].Value, // Array index for description to be corrected as per VTFK rest return
                  Sektor: ""//arr[16 + this.countArrIncrease].Value.split(';')[1].split('|')[2]
              }
            ); 
          }
        });
      }
    }
    else
    {
      this.subNews = [];
      
      if (items !== undefined && items.length > 0) {
        items.forEach((list:any)=>{  
          // Loads the main news details  
         if (!isHubItems) {
          this.subNews.push(
            {
                Title:list.Title === undefined || list.Title === null ? "":list.Title, 
                Created: list.ListItemAllFields === undefined ? "" : list.ListItemAllFields.Modified,
                Author: list.Author === undefined ? "" : list.Author.Title,
                Name: list.Name === undefined || list.Name === null ? "" : list.Name.replace(".aspx", ""),
                ViewUrl: list.ServerRelativeUrl === undefined ? 
                          list.AbsoluteUrl === undefined ? "#" : 
                          list.AbsoluteUrl : 
                          this.GetDominUrl() + list.ServerRelativeUrl,
                Description: list.ListItemAllFields.Description === undefined || list.ListItemAllFields.Description === null ? "":list.ListItemAllFields.Description,
                Sektor: list.ListItemAllFields.Sektor0 === undefined || list.ListItemAllFields.Sektor0 === null ? "":list.ListItemAllFields.Sektor0.Label,
                UniqueID: list.UniqueId === undefined || list.UniqueId === null ? "" : list.UniqueId
            }
          ); 
         }
         else
         {
          var arr = list.Cells;

          this.subNews.push(
            {
              Title: arr[1 + this.countArrIncrease].Value, 
              Created: arr[2 + this.countArrIncrease].Value,
              Author: arr[4 + this.countArrIncrease].Value !== null ? arr[4 + this.countArrIncrease].Value.split('|')[1]:"-",
              AuthorEmail: arr[4 + this.countArrIncrease].Value !== null ? arr[4 + this.countArrIncrease].Value.split('|')[0 + this.countArrIncrease]:"noemail",
              ViewUrl: arr[5 + this.countArrIncrease].Value,
              SiteUrl: arr[6 + this.countArrIncrease].Value,
              Name: arr[5 + this.countArrIncrease].Value !== null ? arr[5 + this.countArrIncrease].Value.substring(arr[5 + this.countArrIncrease].Value.lastIndexOf('/') + 1).replace(".aspx","") : "",
              Description: arr[7 + this.countArrIncrease].Value, // Array index for description to be corrected as per VTFK rest return
             // Sektor: arr[8 + this.countArrIncrease].Value !== null ? arr[8 + this.countArrIncrease].Value.contains(";") 
                 //   ? arr[8 + this.countArrIncrease].Value.split(';')[1].split('|')[2] 
                 //   : arr[8 + this.countArrIncrease].Value :"",
              UniqueID: arr[0 + this.countArrIncrease].Value !== null ? arr[0 + this.countArrIncrease].Value:"",
              Sektor: arr[8 + this.countArrIncrease].Value !== null ? arr[8 + this.countArrIncrease].Value : ""
              //arr[16 + this.countArrIncrease].Value.split(';')[1].split('|')[2]
            }
          ); 
         }
        });
      }
    }
    
    this.requireToRenderHTML = true;
    this.render();
  }

  private GetJsonData(listresturl:string): Promise<any>{  
    // Call to site to get the library names/ sites for hub  
    const res = this.context.httpClient
    .get(listresturl, SPHttpClient.configurations.v1,
      {
        headers: [
          ['accept', 'application/json']
        ]
      })
    .then((res: SPHttpClientResponse) => {
      return res.json();
    });

    return Promise.resolve<any>(res);
  }

  private LoadPropertyDropDownData(listresturl:string): Promise<propertyFields>{  
    // Call to site to get the library names/ sites for hub  
    const res = this.context.httpClient
    .get(listresturl, SPHttpClient.configurations.v1,
      {
        headers: [
          ['accept', 'application/json']
        ]
      })
    .then((res: SPHttpClientResponse) => {
      return res.json();
    });

    return Promise.resolve<any>(res);
  }

  private LoadLibraryOrSiteDropDownData(listresturl:string): Promise<spLists>{  
    // Call to site to get the library names/ sites for hub  
    const res = this.context.httpClient
    .get(listresturl, SPHttpClient.configurations.v1,
      {
        headers: [
          ['accept', 'application/json']
        ]
      })
    .then((res: SPHttpClientResponse) => {
      return res.json();
    });

    return Promise.resolve<any>(res);
  }

  private GetFilesFromLibOrHub(): void{  
    // Retrives Items from SP lib/ hub  
    if(this.properties.sourceLib !== undefined ||
      (this.properties.sourceLib === undefined &&
       this.properties.newsSiteUrl !== undefined &&
       this.properties.newsSiteUrl.length > 20 )){ 
      
      let url: string = "";

      if(this.properties.sourceSite === "site")
      {
        url = this.context.pageContext.web.absoluteUrl + 
                        "/_api/web/lists/getbytitle('" + this.properties.sourceLib + 
                        "')/items?filter=FileSystemObjectType eq 0&$select=FileLeafRef,File/ServerRelativeUrl&$expand=File&$top=2000";
      }
      else if(this.properties.sourceSite === "hub")
      {
        url = this.context.pageContext.web.absoluteUrl +
              "/_api/search/query?querytext='IsDocument:True AND DepartmentId:{" + this.properties.sourceLib + "} AND FileExtension:aspx'&rowlimit=5000";
        // AND PromotedState:2
      }
      else if(this.properties.newsSiteUrl !== undefined &&
              (this.properties.sourceSite === "news" || 
              (this.properties.sourceSite === undefined &&
              this.properties.newsSiteUrl.length > 20)
              ))
      {
        url = this.properties.newsSiteUrl + 
        "/_api/web/lists/getbytitle('" + this.properties.sourceLib + 
        "')/items?filter=FileSystemObjectType eq 0&$select=UniqueId,FileLeafRef,File/ServerRelativeUrl&$expand=File&$top=2000";
      }

      if (url !== "") {
        console.log("Pages: " + url);
        this.GetPageFromLibOrHubSite(url).then((response)=>{  
          // Loads in to drop down field
          var retRes = response.value !== undefined ? response.value : response;
          this.PopulatePagesDropDown(retRes);  
        });  
      }
      else
      {
        this.pageDropDownOptions = [];
        this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
      }
    }  
  }  
    
  private GetPageFromLibOrHubSite(listresturl:string): Promise<spListItems>{  

    // Cal to get page from lib or Hub
    const res = this.context.httpClient
    .get(listresturl, SPHttpClient.configurations.v1,
      {
        headers: [
          ['accept', 'application/json']
        ]
      })
    .then((res: SPHttpClientResponse) => {
      return res.json();
    });

    return Promise.resolve<any>(res);
  }  
    
  private PopulatePagesDropDown(listitems: any): void{  
    // Populates pages drop down values  
    this.pageDropDownOptions = []; 

    if(listitems != undefined)
    {  
      if(this.properties.sourceSite === "site" ||
         this.properties.sourceSite === "news" || 
         (this.properties.sourceSite === undefined &&
          this.properties.newsSiteUrl !== undefined &&
          this.properties.newsSiteUrl.length > 20))
      {
        listitems.forEach((listItem:spListItem)=>{  
          this.pageDropDownOptions.push({key: listItem.FileLeafRef,text:listItem.FileLeafRef});  
        });  
      }
      else if(this.properties.sourceSite === "hub")
      {
        if (listitems.PrimaryQueryResult.RelevantResults.Table.Rows.length > 0) 
        {
          listitems = listitems.PrimaryQueryResult.RelevantResults.Table.Rows;
          listitems.forEach((listItem)=>{
            this.pageDropDownOptions.push({key:listItem.Cells[6].Value,text:this.GetPageNameFromUrl(listItem.Cells[6].Value)});  //Cells[5], Cells[2] for docsnode
          }); 
        }
      }
    }
    
    this.properties.isAdvSettingEnabled !== undefined ? this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled):"";
  } 

  private ClearPropertyFields(): void
  {
    this.properties.sourceLib = "";
    this.properties.sourcePage = "";
    this.properties.filterCondition = "";
  }

  private GetDominUrl()
  {
    var arr = window.location.href.split("/");
    var domainUrl = arr[0] + "//" + arr[2];

    return domainUrl;
  }

  private AddJSFileToHeader(jsFileUrl: string)
  {
    // DOM: Create the script element
    var jsElm = document.createElement("script");
    // set the type attribute
    jsElm.type = "application/javascript";
    // make the script element load file
    jsElm.src = jsFileUrl;
    // finally insert the element to the body element in order to load the script
    document.body.appendChild(jsElm);

    console.log("added: DateTime Js");
  }

  private GetPageGuid(siteType: string, 
    listresturlAA: string, domainUrlAA: string, 
    fileNameAA: string): Promise<any>{ 

	  var pageGUIDURL = listresturlAA + "/_api/web/getfolderbyserverrelativeurl('" + 
                      listresturlAA.replace(domainUrlAA,"") + 
                      "/SitePages')/Files?$expand=ListItemAllFields&$filter=Name eq '" + 
                       fileNameAA + "'&$select=UniqueId";
 
  	const res = this.context.httpClient
    .get(pageGUIDURL, SPHttpClient.configurations.v1,
      {
        headers: [
          ['accept', 'application/json']
        ]
      })
    .then((res: SPHttpClientResponse) => {
      return res.json();
    });

    return Promise.resolve<any>(res);
}
  private GetNewsUrl(newsType: string, siteType: string, 
                     listresturl: string, domainUrl: string, 
                     fileName: string)
  {
    var retUrl = "";

    if (newsType === "main") 
    {
      if (siteType === "hub") 
      {
        retUrl = this.context.pageContext.web.absoluteUrl + "/_api/search/query?querytext='Path:" + this.properties.sourcePage + 
                 "'&selectproperties='Title,LastModifiedTime,Author,AuthorOWSUSER,Path,SiteName,Name,Description,owstaxIdSektor0'"; 
      }
      else
      {
        retUrl = listresturl + "/_api/web/getfolderbyserverrelativeurl('" + 
                              listresturl.replace(domainUrl,"") + 
                              "/SitePages')/Files?$expand=ListItemAllFields&$filter=Name eq '" + 
                              fileName + "'&$select=*,Author/Title,Author/Email&$expand=Author/ID";
      }
    }
    else if(newsType === "sub")
    {
      if (siteType === "hub") 
      {
        retUrl = this.context.pageContext.web.absoluteUrl + "/_api/search/query?querytext='IsDocument:True AND DepartmentId:{"+ this.properties.sourceLib +"} AND FileExtension:aspx'&rowlimit=4&sortlist='ValoNewsPublishDate:descending'&selectproperties='Title,LastModifiedTime,Author,AuthorOWSUSER,Path,SiteName,Name,Description,Sektor,ValoNewsPublishDate'";
      }
      else
      {
        if (this.properties.filterCondition.length > 0 && this.properties.filterCondition.charAt(0) == "'") 
        {
          retUrl = listresturl + '/_api/search/query?queryTemplate=' + this.properties.filterCondition + 
                   "&rowlimit=4&sortlist='ValoNewsPublishDate:descending'&selectproperties='uniqueid,Title,LastModifiedTime,Author,AuthorOWSUSER,Path,SiteName,Name,Description,owstaxIdSektor0,ValoNewsPublishDate'";
        }
        else
        {
          retUrl = listresturl + "/_api/web/getfolderbyserverrelativeurl('" + 
                              listresturl.replace(domainUrl,"") + 
                              "/SitePages')/Files?$expand=ListItemAllFields" + 
                                "&$select=*,Author/Title&$expand=Author/ID&$top=4&$orderby=TimeLastModified desc"+
                                "&$filter=substringof('.0',UIVersionLabel)";
        }
      }
    }

    return retUrl;
  }

  private ButtonClickAdd(oldVal: any): any 
  { 
    if(this.properties.searchValue !== undefined && 
       this.properties.searchValue != "" && 
       this.properties.filterValue !== undefined &&
       this.properties.filterValue != "")
    {
        var opr = this.properties.operatorType === undefined ? "eq" : this.properties.operatorType;

        if(this.properties.filterCondition !== undefined && 
          this.properties.filterCondition.length > 5)
        {
          this.properties.filterCondition = "(" + this.properties.filterCondition + ")";
          this.properties.filterCondition += " and "

          if (opr == "eq") {
            this.properties.filterCondition += "(ListItemAllFields/" + this.properties.searchValue + " " + opr + 
                                            " '" + this.properties.filterValue + "')";
          }
          else
          {
            this.properties.filterCondition += "(substringof('" + this.properties.filterValue + 
                                               "',ListItemAllFields/" + this.properties.searchValue  + "))"
          }
        }
        else
        {
          this.properties.filterCondition = "";

          if (opr == "eq") {
            this.properties.filterCondition += "ListItemAllFields/" + this.properties.searchValue + " " + opr + 
                                            " '" + this.properties.filterValue + "'";
          }
          else
          {
            this.properties.filterCondition += "substringof('" + this.properties.filterValue + 
                                               "',ListItemAllFields/" + this.properties.searchValue  + ")"
          }
        }

        this.properties.searchValue ="";
        this.properties.filterValue = "";
        this.properties.operatorType = "eq";
       
        this.requireToLoadNews = true;
    }
    else
    {
      alert("Please provide filter column and value");
    }
  }

  private ButtonClickEdit(oldVal: any): any 
  { 
    this.enableFilterTextbox = false;
    this.GetButtonsForProperty(true);
    this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);
  }

  private ButtonClickSave(oldVal: any): any 
  { 
    this.enableFilterTextbox = true;
    this.GetButtonsForProperty(false);
    this.SetPropertyPaneFields(this.properties.isAdvSettingEnabled);

    this.requireToLoadNews = true;
  }

  private ButtonClickClear(oldVal: any): any 
  { 
    this.properties.filterCondition = "";
    
    this.requireToLoadNews = true;
  }

  private _onSearch(value: string)
  {
    
  }

  private GetPageNameFromUrl(url: string)
  {
    return url !== undefined && url !== null ? url.substring(url.lastIndexOf('/') + 1).replace(".aspx","") : "";
  }

  private GetButtonsForProperty(isEditMode: boolean)
  {
    this.propFilterButtons = [];
    
    if(!isEditMode)
    {
      this.propFilterButtons.push(
                                    PropertyPaneButton('addFilter',  
                                      {  
                                        text: "",  
                                        buttonType: PropertyPaneButtonType.Hero,  
                                        onClick: this.ButtonClickAdd.bind(this),
                                        icon:'Add'  
                                      })
                                  );
      this.propFilterButtons.push(
                                    PropertyPaneButton('editFilter',  
                                      {  
                                        text: "",  
                                        buttonType: PropertyPaneButtonType.Hero,  
                                        onClick: this.ButtonClickEdit.bind(this),
                                        icon:'edit'
                                      })
                                  );
      this.propFilterButtons.push(
                                    PropertyPaneButton('clearFilter',  
                                      {  
                                        text: "",  
                                        buttonType: PropertyPaneButtonType.Hero,  
                                        onClick: this.ButtonClickClear.bind(this),
                                        icon:'clear'
                                      })
                                  );
    }
    else
    {
        this.propFilterButtons.push(
                  PropertyPaneButton('saveFilter',  
                    {  
                      text: "Save",  
                      buttonType: PropertyPaneButtonType.Hero,  
                      onClick: this.ButtonClickSave.bind(this),
                      icon:'save'
                    })
                );
    }
  }

  private SetPropertyPaneFields(isAdvMode: boolean)
  {
    this.propPaneFields = [];

    this.propPaneFields.push(
                              PropertyPaneToggle('showAuthor',{
                                label:"Author:",
                                checked: true
                              }),
                              PropertyPaneToggle('showDate',{
                                label:"Date:",
                                checked: true
                              }),
                              PropertyPaneToggle('requireTitleTruncate',{
                                label:"Title to be truncated:",
                                checked: false
                              }),
                              PropertyPaneToggle('requireDescTruncate',{
                                label:"Description to be truncated:",
                                checked: false
                              }),
                              PropertyPaneToggle('isAdvSettingEnabled',{
                                label:"Advance Settings:"
                              })
                            );
    if (isAdvMode !== undefined &&
        isAdvMode) 
    {
      this.propPaneFields.push(
                                PropertyPaneChoiceGroup('sourceSite', {
                                  options:[{key:'news', text:'News Site', checked: true},{key:'site', text:'Current Site'},{key:'hub', text:'Hub Site'}],
                                  label: strings.SourceSiteLabel
                                }),
                                PropertyPaneDropdown('sourceLib', {
                                  options:this.libOrSiteDropDownOptions,
                                  label: this.properties.sourceSite === "hub" ? "Hub Sites" : strings.SourceLibLabel,
                                  disabled: this.properties.sourceSite !== "hub",
                                  selectedKey: this.properties.sourceSite === "hub" ? "" : "Site Pages"
                                }),
                                PropertyFieldSearch("sourcePage", {
                                  key: "search",
                                  placeholder: 'Select Campaign Page',
                                  value: this.properties.sourcePage,
                                  styles: { root: { margin: 10 } }
                                }),
                                PropertyPaneChoiceGroup('selectedPage', {
                                  options:this.selectedPages,
                                  label: ""
                                }),
                                PropertyPaneTextField('newsSiteUrl', {
                                  label: "Configure News site:",
                                  disabled: this.properties.sourceSite !== "news"
                                }),
                                PropertyPaneTextField('filterCondition', {
                                  multiline:true,
                                  disabled: this.enableFilterTextbox
                                }),
                                PropertyFieldSearch("searchValue", {
                                  key: "search",
                                  placeholder: 'Property',
                                  value: this.properties.searchValue,
                                  onSearch: this._onSearch,
                                  styles: { root: { margin: 10 } }
                                }),
                                PropertyPaneChoiceGroup('selectedValue', {
                                  options:this.selectedFields,
                                  label: ""
                                }),
                                PropertyPaneDropdown('operatorType', {
                                  options:[{key:'eq', text:'Equal'},{key:'contains', text:'Contains'}],
                                  label: "",
                                  selectedKey:"eq"
                                }),
                                PropertyPaneTextField('filterValue', {
                                  placeholder:"Value"
                                })
                          );
    }

    this.context.propertyPane.refresh();
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: this.propPaneFields.concat(this.propFilterButtons)
            }
          ],
        }
      ]
    };
  }
}
