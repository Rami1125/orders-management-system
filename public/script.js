function sendData() {
  const data = { message: "שלום מגיטהאב!" };
  google.script.run.withSuccessHandler(response => {
    alert("נשלח: " + response);
  }).doPost({ postData: JSON.stringify(data) });
}

