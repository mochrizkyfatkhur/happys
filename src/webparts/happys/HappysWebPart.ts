import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { PropertyPaneButton } from '@microsoft/sp-webpart-base';
//import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HappysWebPart.module.scss';

import * as strings from 'HappysWebPartStrings';

export interface IHappysWebPartProps {
  description: string;
}


import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';


import { SPComponentLoader } from '@microsoft/sp-loader';


export interface ISPLists {
  value: ISPList[];
}


export interface ISPList {
  Title: string;
  fxday: string;
  fxmonth: string;
  fxuser:string;
}


export interface IHappysWebPartProps {
  url: string;
  description: string;

  //try
  employeename: string;
  birthday: string;
}


export default class HappysWebPart extends BaseClientSideWebPart<IHappysWebPartProps> {


  public render(): void {
    var titlemonth= this.getmonth();

 

    this.domElement.innerHTML = `

    <div class="${styles.happys}">

      <div class="container">

      <br>

          <div class="card">

          <h1 class="text-center ${styles.titles} ${styles.tada} ${styles.animated} ${styles.infinite}" style="color: black;"><img src="https://bl3301files.storage.live.com/y4m4lxYslyp1ijch7Yf2ohzKio492J46gzEC05PFyiPUzPur4zXUsbET5rhR7laS1jka3pJ2I3wQOk9iRpgiVsz2bcbHSIOIlsAwHK2fjbqMkrcHrFWH5TTj0pPY9sqmectoz5ty-Hp-adJepOG_UDh2WzT77cqEdh0BzX9sDAfSF1Mkpkek8rfE3u_A-oMozw_?width=512&height=512&cropmode=none" alt="" style="width: 40px;height: auto;"/>Happy birthday<img src="https://bl3301files.storage.live.com/y4m4lxYslyp1ijch7Yf2ohzKio492J46gzEC05PFyiPUzPur4zXUsbET5rhR7laS1jka3pJ2I3wQOk9iRpgiVsz2bcbHSIOIlsAwHK2fjbqMkrcHrFWH5TTj0pPY9sqmectoz5ty-Hp-adJepOG_UDh2WzT77cqEdh0BzX9sDAfSF1Mkpkek8rfE3u_A-oMozw_?width=512&height=512&cropmode=none" alt="" style="width: 40px;height: auto;"/></h1>

                   

          <p class="text-center  " style="color: black;">Birthday List of the Month</p>

          <img src="https://bl3301files.storage.live.com/y4m0WhAwqEfh_ON1ooWD4ZIBv0RYFKegOMYQ8FpMFP0sEvPeNUBoxOE1g-TEIpnF5SozjCDsqgQlBFSDAzAFf7nORy07yS3Dtu3ReZh5xrD-p6qwTE9QyRJF9VxYPx79sqmDbSrexnJHbdx2w5pPOvR7zqKLNonDNVa4VBBI_Vb1vmECrG3TF8NbSDqeBpNBLdm?width=512&height=512&cropmode=none" class="rounded mx-auto d-block ${styles.pulse} ${styles.animated} ${styles.infinite}" style="width: 120px;height: auto;">

          <br>

          <h2 class="text-center ${styles.titles}" style="color: black;">${titlemonth}</h2>

          <br>

          <table class="table">

    <thead>

      <tr>

        <th class="text-center">Name</th>

        <th class="text-center">Date</th>

        <th class="text-center">Email</th>

      </tr>

    </thead>

    <tbody id="spList">



    

    </tbody>

  </table>

          </div>

        </div>

        </div>`;

        this._renderListAsync();
  }


  protected getmonth(){
    var month = new Array();
    month[0] = "January";
    month[1] = "February";
    month[2] = "March";
    month[3] = "April";
    month[4] = "May";
    month[5] = "June";
    month[6] = "July";
    month[7] = "August";
    month[8] = "September";
    month[9] = "October";
    month[10] = "November";
    month[11] = "December";


    var d = new Date();
    var n = month[d.getMonth()];
    return n;
  }


  protected getnumber(){
    var month = new Array();
    month[0] = 1;
    month[1] = 2;
    month[2] = 3;
    month[3] = 4;
    month[4] = 5;
    month[5] = 6;
    month[6] = 7;
    month[7] = 8;
    month[8] = 9;
    month[9] = 10;
    month[10] = 11;
    month[11] = 12;


    var d = new Date();
    var n = month[d.getMonth()];
    return n;
  }


  public constructor() {
    super();
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://fonts.googleapis.com/css2?family=Dancing+Script&display=swap');
  }


  


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  private _getListData(): Promise<ISPLists> {

    return this.context.spHttpClient.get(this.properties.url + `/_api/web/lists/GetByTitle('birthdays')/Items`, SPHttpClient.configurations.v1)

      .then((response: SPHttpClientResponse) => {

 

        return response.json();

      });

  }
  
  private _renderList(items: ISPList[]): void {

    let html: string = '';

    items.forEach((item: ISPList) => {

      console.log(item.fxmonth);

      console.log(this.getnumber());

      if(item.fxmonth==this.getnumber()){

      html += `

    <tr>

    <td class="text-center h6 ${styles.titles}">${item.Title}</td>

    <td class="text-center h6 ${styles.titles}">${item.fxday}</td>

    <td class="text-center h6 ${styles.titles}">${item.fxuser}</td>
    </tr>`;

      }

    });

 

    const listContainer: Element | null = this.domElement.querySelector('#spList');

    if (listContainer) {
        listContainer.innerHTML = html;
    }

  }



  private _renderListAsync(): void {
      
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
  
}



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Masukan Url untuk mengakses data dari SharePoint"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('url', {
                  label: "Urls",
                  //value: "https://jababekainfrastruktur.sharepoint.com/:l:/s/birthdays/FIFB35w4RJ1CvoW-TnftwnUBLoNDfwQBgsY7w3hfxAc65w?e=MDUg1p"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}





// import { Version } from '@microsoft/sp-core-library';
// import {
//   IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-property-pane';
// import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// //import { PropertyPaneButton } from '@microsoft/sp-webpart-base';
// //import { escape } from '@microsoft/sp-lodash-subset';

// import styles from './HappysWebPart.module.scss';

// import * as strings from 'HappysWebPartStrings';

// export interface IHappysWebPartProps {
//   description: string;
//   url: string;
// }

// // import {
// //   SPHttpClient,
// //   SPHttpClientResponse
// // } from '@microsoft/sp-http';

// import { SPComponentLoader } from '@microsoft/sp-loader';

// export interface ISPLists {
//   value: ISPList[];
// }

// export interface ISPList {
//   Title: string;
//   fxday: string;
//   fxmonth: number;
//   fxuser: string;
// }

// import * as XLSX from 'xlsx';

// export default class HappysWebPart extends BaseClientSideWebPart<IHappysWebPartProps> {

//   public render(): void {
//     var titlemonth = this.getmonth();

//     this.domElement.innerHTML = `
//       <div class="${styles.happys}">
//         <div class="container">
//           <br>
//           <div class="card">
//             <h1 class="text-center ${styles.titles} ${styles.tada} ${styles.animated} ${styles.infinite}">
//               <img src="https://bl3301files.storage.live.com/y4m4lxYslyp1ijch7Yf2ohzKio492J46gzEC05PFyiPUzPur4zXUsbET5rhR7laS1jka3pJ2I3wQOk9iRpgiVsz2bcbHSIOIlsAwHK2fjbqMkrcHrFWH5TTj0pPY9sqmectoz5ty-Hp-adJepOG_UDh2WzT77cqEdh0BzX9sDAfSF1Mkpkek8rfE3u_A-oMozw_?width=512&height=512&cropmode=none" alt="" style="width: 40px;height: auto;"/>
//               Happy birthday
//               <img src="https://bl3301files.storage.live.com/y4m4lxYslyp1ijch7Yf2ohzKio492J46gzEC05PFyiPUzPur4zXUsbET5rhR7laS1jka3pJ2I3wQOk9iRpgiVsz2bcbHSIOIlsAwHK2fjbqMkrcHrFWH5TTj0pPY9sqmectoz5ty-Hp-adJepOG_UDh2WzT77cqEdh0BzX9sDAfSF1Mkpkek8rfE3u_A-oMozw_?width=512&height=512&cropmode=none" alt="" style="width: 40px;height: auto;"/>
//             </h1>
//             <p class="text-center">Birthday List of the Month</p>
//             <img src="https://bl3301files.storage.live.com/y4m0WhAwqEfh_ON1ooWD4ZIBv0RYFKegOMYQ8FpMFP0sEvPeNUBoxOE1g-TEIpnF5SozjCDsqgQlBFSDAzAFf7nORy07yS3Dtu3ReZh5xrD-p6qwTE9QyRJF9VxYPx79sqmDbSrexnJHbdx2w5pPOvR7zqKLNonDNVa4VBBI_Vb1vmECrG3TF8NbSDqeBpNBLdm?width=512&height=512&cropmode=none" class="rounded mx-auto d-block ${styles.pulse} ${styles.animated} ${styles.infinite}" style="width: 120px;height: auto;">
//             <br>
//             <h2 class="text-center ${styles.titles}">${titlemonth}</h2>
//             <br>
//             <input type="file" id="fileUpload" accept=".xlsx, .xls" />
//             <br><br>
//             <table class="table">
//               <thead>
//                 <tr>
//                   <th class="text-center">Name</th>
//                   <th class="text-center">Day</th>
//                   <th class="text-center">Chat</th>
//                 </tr>
//               </thead>
//               <tbody id="spList"></tbody>
//             </table>
//           </div>
//         </div>
//       </div>`;
//     this._attachFileUploadEvent();
//   }

//   private _attachFileUploadEvent(): void {
//     const fileUpload: HTMLInputElement = this.domElement.querySelector('#fileUpload') as HTMLInputElement;
//     fileUpload.addEventListener('change', (event: Event) => this._handleFileUpload(event));
//   }

//   private _handleFileUpload(event: Event): void {
//     const input = event.target as HTMLInputElement;
//     if (input && input.files && input.files[0]) {
//       const file = input.files[0];
//       const reader = new FileReader();
//       reader.onload = (e: ProgressEvent<FileReader>) => {
//         const data = new Uint8Array(e?.target?.result as ArrayBuffer);
//         const workbook = XLSX.read(data, { type: 'array' });
//         const firstSheetName = workbook.SheetNames[0];
//         const worksheet = workbook.Sheets[firstSheetName];
//         const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
//         this._renderListFromExcel(jsonData);
//       };
//       reader.readAsArrayBuffer(file);
//     }
//   }

//   private _renderListFromExcel(data: any[]): void {
//     const currentMonth = this.getnumber();
//     let html: string = '';
//     data.forEach((row, index) => {
//       if (index === 0) return; // Skip header row
//       const [name, day, month, email] = row;
//       if (month == currentMonth) {
//         html += `
//           <tr>
//             <td class="text-center h5 ${styles.titles}">${name}</td>
//             <td class="text-center h5 ${styles.titles}">${day}</td>
//             <td class="text-center h5 ${styles.titles}">${email}</td>
//           </tr>`;
//       }
//     });

//     const listContainer: Element | null = this.domElement.querySelector('#spList');
//     if (listContainer) {
//       listContainer.innerHTML = html;
//     }
//   }

//   protected getmonth() {
//     var month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
//     var d = new Date();
//     return month[d.getMonth()];
//   }

//   protected getnumber() {
//     var d = new Date();
//     return d.getMonth() + 1;
//   }

//   public constructor() {
//     super();
//     SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css');
//     SPComponentLoader.loadCss('https://fonts.googleapis.com/css2?family=Dancing+Script&display=swap');
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: "Masukan Url untuk mengakses data dari SharePoint"
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('url', {
//                   label: "Urls"
//                 })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }