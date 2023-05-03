
Office.onReady();

function onAppointmentSendHandlerJS(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Failed to send.",
  });
  return;
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
//if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  console.log("99");
  //Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandlerJS);
  Office.actions.associate("onMessageSendHandler", onAppointmentSendHandlerJS);
//}
