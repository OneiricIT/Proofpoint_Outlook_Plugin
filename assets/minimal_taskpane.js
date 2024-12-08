Office.onReady(() => {
    console.log("Office.js is ready.");
  });
  
  function onMinimalButtonClick(event) {
    console.log("Minimal button clicked.");
    event.completed();
  }
  
