var scElement;
var wtElement;
var okBtn;

Office.onReady(() => {
    let params = new URLSearchParams(window.location.search);
    let sColor = "#"+params.get('color');
    let target = params.get('target');
    
    scElement = document.getElementById("sectionColor");
    wtElement = document.getElementById("wordTargetCount");
    scElement.value = sColor;
    wtElement.value = target;

    document.addEventListener("keypress", function(event) { if (event.key === "Enter") {
        document.getElementById("okBtn").click();
        }
      });
  });

async function confirmChanges(){
    Office.context.ui.messageParent(JSON.stringify([scElement.value,wtElement.value]));
}

function closeDialog(){
    Office.context.ui.messageParent("");
}
