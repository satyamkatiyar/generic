/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// //is this EVN based?
// const applicationId = "Nasdaq.Office.Service.Daemon.V1";
// const secretKey = "Mk19WrGtFMWecdVi";
// const PRODUCT_ID = 1300;

// Cookies.set('dragonTicket1', localStorage.getItem("tempDataCopy"));

// Office.onReady();
// //debugger;

// //saveAnEvent
// //uploadAttachments
// //saveAttachments
// //updateEventInCalendar

// function onItemSendHandler3(event) {
//     debugger;
//     console.log("**********onItemSendHandler3");
    
//     Office.context.mailbox.item.categories.getAsync({ asyncContext: event },
//         (asyncResult) => {
//         console.log("**********categories.getAsync");
//         console.log(asyncResult);
//         if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//             const categories = asyncResult.value;
//           if (categories && categories.length > 0) {
//             if (
//               categories.some((r) => r.displayName.includes("Nasdaq CMS Meetings"))
//             ) {
//               event.completed({ allowEvent: false,
//                 errorMessage: "Looks like add-in is down right now." });
//               return;
//             CreateCMSRequest(event);
            

//             //CALL NDSService.saveAnEvent
//             //CALL updateEventInCalendar --> UpdateCalendarEventWithId
//             }
//           } else {
//             console.log("There are no categories assigned to this item.");
//           }
//         } else {
//           console.error("Failed to get categorie",asyncResult.error);
//         }
//       });
//   }

//   function CreateCMSRequest(event){
//     let mailboxItem = Office.context.mailbox.item;
//         mailboxItem.subject.getAsync({ asyncContext: event }, (asyncResult) => {
//           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//             console.log("asyncResult.error.message");
//             console.log(asyncResult.error.message);
//           } else {
//             let request = JSON.parse(localStorage.getItem("CMSRequest"));// ApplicationStorage.getCMSRequest();
//             // Successfully got the subject, display it.
//             request.eventTitle = asyncResult.value;
//             console.log(request.eventTitle);
//             //ApplicationStorage.setItem("CMSRequest", JSON.stringify(request));
//             localStorage.setItem("CMSRequest", JSON.stringify(request));
//             request = JSON.parse(localStorage.getItem("CMSRequest"));
//             console.log("REQUEST GOING", request);
//             SaveEventX(event, request);
//           }
//         });
//   }

//   function axiosInterceptor(){
//     // Add a request interceptor
//     axiosinstance.interceptors.request.use(
//       function (config) {
//           // update common request settings!
//           debugger;
//           config.withCredentials = true;
//           if (!config.url.includes("api/Dataitems")) {
//               if (config.method === "post") {
//                 if(!config.headers.post){
//                   config.headers.post = {};
//                   config.headers.post['Content-Type'] = 'application/json';
//                   config.headers.post['Accept'] = "application/json";
//                 }
//                 getSecurityHeadersX(config.url, config.headers.post);
//               } else if (config.method === "delete") {
//                 getSecurityHeadersX(config.url, config.headers.delete);
//               } else if (config.method === "put") {
//                 getSecurityHeadersX(config.url, config.headers.put);
//               }
//           }
//           return config;
//       },
//       function (error) {
//           // Do something with request error
//           return Promise.reject(error);
//       }
//     );
//   }

//   function SaveEventX(event, request){
//         console.log("now we will call save event api");
//         console.log(request);
//         let apiurl = "https://irinsight-devint.nasdaq.com/DataServices/api" + "/meetings/SaveEvent"
//         let axiosHeaders = getAxiosHeader(apiurl);
//         let axiosinstance = axios.create({
//           headers: axiosHeaders
//           , withCredentials: true
//         });
//         try {
//           const response = axiosinstance.post(apiurl
//             //,{ withCredentials: true }
//             ,JSON.stringify(request)
//             // ,{
//             //   headers: axiosHeaders
//             // }
//           )
//           .then(response => {
//             if (response.status === 200) {
//               console.log("NDS save event", response);
//               //return response.data;
//               event.completed({ allowEvent: true });
//             }
//             })
//           .catch(error => {
//             console.error(error)
//             event.completed({ allowEvent: false,
//               errorMessage: "Looks like add-in is down right now." });
//           });
//           }
          
//         catch (error) {
//           console.error(error);
//           //this.throwError(error);
//         }
//         //https://irinsight-devint.nasdaq.com/DataServices/api

//   }



//   let timezoneRegex = /\(([^)]+)\)/;

//   function getFormattedDate (dateValue) {
//     let timezoneID = timezoneRegex.exec(dateValue)[1];
//     let dateISO = new Date(dateValue).toISOString();
//     var formattedDateISO = new Date(dateISO);
//     var formattedDateTime = new Date(
//       formattedDateISO.getTime() - formattedDateISO.getTimezoneOffset() * 60000
//     );
//     let formattedDateTimeIS0 = new Date(formattedDateTime).toISOString();
//     return {
//       timezoneID: timezoneID,
//       formattedDate: formattedDateTimeIS0.substring(0, 19),
//     };
//   }

//   async function updateEventInCalendar (EventId, event) {
//     Office.context.mailbox.item.saveAsync(async function (result) {
//       if (result.status !== Office.AsyncResultStatus.Succeeded) {
//         console.log("Failed to update the event");
//         return null;
//       } else {
//         console.log("ItemId", result.value);
//         //MS Graph API ???
//         // await UpdateCalendarEventWithId(
//         //   localStorage.getItem("AccessToken"),
//         //   result.value,
//         //   EventId,
//         //   event
//         // ).then((res) => {
//         //   // if(res.status)
//         // });
//       }
//     });
//   }

// console.log("**********inlauncheventv2");

// //#region - axio configuration


// function getAxiosHeader(apiUrl){
  
//   let authnonce = getNonce();
//   let authtimestamp = getTimeStamp();
//   let axiosheader = {};
//   axiosheader['Accept'] = "application/json";
//   axiosheader['Content-Type'] = "application/json";
//   axiosheader['Access-Control-Allow-Credentials'] = true;
//   axiosheader['Access-Control-Allow-Origin'] = "*";
//   axiosheader['ProductId'] = PRODUCT_ID;
//   axiosheader['auth_application_id'] = applicationId;
//   axiosheader['auth_nonce'] = authnonce;
//   axiosheader['auth_signature'] = getSignature(authnonce, authtimestamp, apiUrl);
//   axiosheader['auth_timestamp'] = authtimestamp;
//   axiosheader['OriginId'] = crypto.randomUUID();
//   //axiosheader['Cookie'] = localStorage.getItem("tempDataCopy");
//   return axiosheader;
// }

// function configureAxios(){
//   //is this EVN based?
//   const applicationId = "Nasdaq.Office.Service.Daemon.V1";
//   const secretKey = "Mk19WrGtFMWecdVi";
//   //Defaults
//   axios.defaults.headers.post['Content-Type'] = 'application/json';
//   axios.defaults.headers.post['Accept'] = "application/json";

//   // Add a request interceptor
//   axios.interceptors.request.use(
//     function (config) {
//         // update common request settings!
//         config.withCredentials = true;
//         if (!config.url.includes("api/Dataitems")) {
//             if (config.method === "post") {
//               if(!config.headers.post){
//                 config.headers.post['Content-Type'] = 'application/json';
//                 config.headers.post['Accept'] = "application/json";
//               }
//               getSecurityHeadersX(config.url, config.headers.post);
//             } else if (config.method === "delete") {
//               getSecurityHeadersX(config.url, config.headers.delete);
//             } else if (config.method === "put") {
//               getSecurityHeadersX(config.url, config.headers.put);
//             }
//         }
//         return config;
//     },
//     function (error) {
//         // Do something with request error
//         return Promise.reject(error);
//     }
//   );



// }

// function getSecurityHeadersX(url, headers){
  
//   //Added for Antiforgery.
//   let authnonce = getNonce();
//   let authtimestamp = getTimeStamp();

//   headers['auth_nonce'] = authnonce;
//   headers['auth_timestamp'] = authtimestamp;
//   headers['auth_application_id'] = applicationId;
//   headers['auth_signature'] = getSignature(authnonce, authtimestamp, url);

//   //Additional Headers
//   const PRODUCT_ID = 1300;
//   headers['ProductId'] = PRODUCT_ID;
//   headers['OriginId'] = crypto.randomUUID();
// }

// const getNonce = () => {
//   return Math.floor(Math.random() * 9000000000) + 1000000000;
// };
// const getTimeStamp = () => {
//   return Math.floor(Date.now() / 1000);
// };
// const getSignature = (authnonce, authtimestamp, url) => {
//   let splitUrl = url.split('.com');
//   let stringBuilder =
//       splitUrl[1] +
//       "&" +
//       "auth_application_id=" +
//       applicationId +
//       "&" +
//       "auth_nonce=" +
//       authnonce +
//       "&" +
//       "auth_timestamp=" +
//       authtimestamp;
//   let signatureBase = encodeURIComponent(stringBuilder).toLowerCase();
//   let signature = cryptohelper(signatureBase);
//   return encodeURIComponent(signature);
// };
// const cryptohelper = (signatureBase) => {
//   let secretKeyHash = CryptoJS.enc.Utf8.parse(secretKey);
//   let signatureBaseHash = CryptoJS.enc.Utf8.parse(signatureBase);
//   let hash = CryptoJS.HmacSHA1(signatureBaseHash, secretKeyHash);
//   let hashInBase64 = CryptoJS.enc.Base64.stringify(hash);
//   return hashInBase64;
// }

// //#endregion - axio configuration




function onItemSendHandler4(event) {
  console.log("**********onItemSendHandler4");
  
  Office.context.mailbox.item.categories.getAsync({ asyncContext: event },
      (asyncResult) => {
      console.log("**********categories.getAsync");
      console.log(asyncResult);
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const categories = asyncResult.value;
        if (categories && categories.length > 0) {
          if (
            categories.some((r) => r.displayName.includes("Nasdaq CMS Meetings"))
          ) {
            event.completed({ allowEvent: false,
              errorMessage: "Looks like add-in is down right now." });
            return;
          }
        } else {
          console.log("There are no categories assigned to this item.");
        }
      } else {
        console.error("Failed to get categorie",asyncResult.error);
      }
    });
}
Office.onReady() 
    .then( function() { 
      console.log("^^^^office is now ready");
      if (Office.context.platform === Office.PlatformType.PC
        || Office.context.platform === Office.PlatformType.Mac
        || Office.context.platform === Office.PlatformType.OfficeOnline) {
      Office.actions.associate("onAppointmentSendHandler", onItemSendHandler4);
    }
    }); 


