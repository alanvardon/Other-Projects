function sendEmail(to,subject, msg){
  GmailApp.sendEmail(to, subject,msg,{htmlBody:msg})  
}

function getEmail(){
  var sheet = getSpreadsheetName("Email Edit");
  SpreadsheetApp.setActiveSheet(sheet)
  var dataRange = sheet.getRange(2,1,1,1)
  data = dataRange.getValues()
  toEmail = data[0][0]
  
  return toEmail
}

function getSpreadsheetName(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var n in sheets){
    if(name==sheets[n].getName()){
      return sheets[n];
    }
  }
  return sheets[0];
}

function prepareSubjectLine(){
  var subjDate = Utilities.formatDate(new Date(), 'Europe/Stockholm', 'dd/MM/yyyy')
  var subjLine = 'Performance Report - ' + subjDate
  return subjLine
}

function prepareReport1(C_Link,Campaign_Name, Client, Delivery_Status, Campaign_Start_Date, Campaign_End_Date, Lead_Goal, Client_Lead_Type, Client_Delivery_Result, _Client_Goal, CPA_Margin, Predicted_Date_to_Goal, Revenue, Last_Received_Lead){
  
//  var curr_sheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = getSpreadsheetName("Email Edit");
  SpreadsheetApp.setActiveSheet(sheet)
  var dataRange = sheet.getRange(2,2,1,1)
  data = dataRange.getValues()
  var msg = data[0][0]
  
  msg = msg.replace(/<CS_Link>/g,C_Link)
  msg = msg.replace(/<Campaign_Name>/g,Campaign_Name)
  msg = msg.replace(/<Client_Name>/g,Client)
  msg = msg.replace(/<Campaign_Delivery_Status>/g,Delivery_Status)
  msg = msg.replace(/<Campaign_Start_Date>/g,Campaign_Start_Date)
  msg = msg.replace(/<Campaign_End_Date>/g,Campaign_End_Date)
  msg = msg.replace(/<Lead_Goal>/g,Lead_Goal)
  msg = msg.replace(/<Client_Lead_Type>/g,Client_Lead_Type)
  msg = msg.replace(/<Client_Delivery_Result>/g,Client_Delivery_Result)
  msg = msg.replace(/<%Client_Goal>/g,_Client_Goal)
  msg = msg.replace(/<CPA_Margin>/g,CPA_Margin)
  msg = msg.replace(/<Estimated_Goal_Date>/g,Predicted_Date_to_Goal)
  msg = msg.replace(/<Total_Revenue>/g,Revenue)
  msg = msg.replace(/<Last_Received_Lead>/g,Last_Received_Lead)
  
  return msg
  //SpreadsheetApp.setActiveSheet(curr_sheet)
}


function prepareReport2(Campaign_Name, Delivery_Status, Campaign_End_Date, Lead_Goal, Client_Lead_Type,_Client_Goal,Avg_Client_Lead_per_day, Avg_Lead_Prediction){
  
  var curr_sheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = getSpreadsheetName("Email Edit");
  SpreadsheetApp.setActiveSheet(sheet)
  var dataRange = sheet.getRange(3,2,1,1)
  data = dataRange.getValues()
  var msg = data[0][0]
  msg = msg.replace(/<Campaign_Name>/g,Campaign_Name)
  msg = msg.replace(/<Campaign_Delivery_Status>/g,Delivery_Status)
  msg = msg.replace(/<Campaign_End_Date>/g,Campaign_End_Date)
  msg = msg.replace(/<Lead_Goal>/g,Lead_Goal)
  msg = msg.replace(/<Client_Lead_Type>/g,Client_Lead_Type)
  msg = msg.replace(/<%Client_Goal>/g,_Client_Goal)
  msg = msg.replace(/<Avg_Client_Lead_per_day>/g,Avg_Client_Lead_per_day)
  msg = msg.replace(/<Avg_Lead_Prediction>/g,Avg_Lead_Prediction)
  return msg

}

function sendReport() {
  
  //sendEmail("alanvardon@gmail.com", "Subjet line", "This is the body")
  
  var sheet = getSpreadsheetName("Campaign Overview")
  var rows = sheet.getLastRow()
  var cols = sheet.getLastColumn()
  var dataRange = sheet.getRange(2,1,rows-1,cols)
  var data  = dataRange.getValues();
  var allreportmsg = ["<!DOCTYPE html><html><body>"]
  var emailR = getEmail()
  var subjLine = prepareSubjectLine()
  
  for (i in data){
    var Client = data[i][0]
    var Campaign_Name = data[i][1]
    var Delivery_Status = data[i][2]
    var Campaign_Type = data[i][3]
    var C_Link = data[i][4]
    var Campaign_Summary_Link = data[i][5]
    var D_Link = data[i][6]
    var Google_Data_Studio_Link = data[i][7]
    var Start_Date = Utilities.formatDate(new Date(data[i][8]),'Europe/Stockholm','dd/MM/yyyy')
    var End_Date = Utilities.formatDate(new Date(data[i][9]),'Europe/Stockholm','dd/MM/yyyy')
    var Last_Received_Lead = Utilities.formatDate(new Date(data[i][10]),'Europe/Stockholm','dd/MM/yyyy')
    var Campaign_Length = data[i][11]
    var Campaign_Elapsed = data[i][12]
    var Days_Left = data[i][13]
    var Lead_Goal = data[i][14]
    var Client_Goal_Type = data[i][15]
    var Client_Delivery_Result = data[i][16]
    var No_Leads = data[i][17]
    var Leads_wPhone = data[i][18]
    var EL = data[i][19]
    var EL_wPhone = data[i][20]
    var Leads_Left = data[i][21]
    var _Client_Goal = (data[i][22]*100).toFixed(2) + "%"
    var Marketing_Spend = data[i][23]
    var Revenue_Goal = data[i][24]
    var CPA = data[i][25]
    var PPL = data[i][26]
    var CPA_Margin = (data[i][27]*100).toFixed(2) + "%"
    var Revenue = parseFloat(data[i][28]).toFixed(2) + " kr"
    var Revenue_Left = data[i][29]
    var _Revenue_Goal = data[i][30]
    var Net = data[i][31]
    var Client_Leads_in_Last_7_Days = data[i][32]
    var Leads_in_Last_7_Days = data[i][33]
    var Leads_wPhone_in_Last_7_Days = data[i][34]
    var EL_in_Last_7_Days = data[i][35]
    var EL_wPhone_in_Last_7_Days = data[i][36]
    var Client_Lead_Ratio = data[i][37]
    var Leads_wPhone_Ratio = data[i][38]
    var EL_Ratio = data[i][39]
    var EL_wPhone_Ratio = data[i][40]
    var Client_Lead_Ratio_for_Last_7_Days = data[i][41]
    var Leads_wPhone_Ratio_7_Days = data[i][42]
    var EL_Ratio_for_Last_7_Days = data[i][43]
    var EL_wPhone_Ratio_for_Last_7_Days = data[i][44]
    var Avg_Client_Lead_per_day = parseFloat(data[i][45]).toFixed(2)
    var Avg_Lead_delivered_per_day = data[i][46]
    var Avg_Lead_wPhone_delivered_per_day = data[i][47]
    var Avg_EL_delivered_per_day = data[i][48]
    var Avg_EL_wPhone_delivered_per_day = data[i][49]
    var Avg_Client_Lead_per_day_7d_MA = data[i][50]
    var Avg_Lead_delivered_per_day_7d_MA = data[i][51]
    var Avg_Lead_wPhone_delivered_per_day_7d_MA = data[i][52]
    var Avg_EL_delivered_per_day_7d_MA = data[i][53]
    var Avg_EL_wPhone_delivered_per_day_7d_MA = data[i][54]
    var Client_Leads_Compltd_Flow = data[i][55]
    var Leads_Compltd_Flow = data[i][56]
    var Leads_wPhone_Compltd_Flow = data[i][57]
    var EL_Compltd_Flow = data[i][58]
    var EL_wPhone_Compltd_Flow = data[i][59]
    var _Client_Leads_Compltd_Flow = data[i][60]
    var _Leads_Compltd_Flow = data[i][61]
    var _Leads_wPhone_Compltd_Flow = data[i][62]
    var _EL_Compltd_Flow = data[i][63]
    var _EL_wPhone_Compltd_Flow = data[i][64]
    var Leads_wInternational_Studies = data[i][65]
    var Predicted_Days_to_Goal = data[i][66]
    var Predicted_Date_to_Goal = Utilities.formatDate(new Date(data[i][67]),'Europe/Stockholm','dd/MM/yyyy')
    var Predicted_Days_to_Goal_7_DMA = data[i][68]
    var Predicted_Date_to_Goal_7d_MA = data[i][69]
    var Avg_Lead_Prediction = (data[i][21]/data[i][13]).toFixed(2)
    if(Math.abs(Avg_Lead_Prediction) == 0 || Math.abs(Avg_Lead_Prediction) == Infinity){Avg_Lead_Prediction = "[No data]"}
    if(Start_Date == "01/01/1970"){Start_Date = "[No data]"}
    if(End_Date == "01/01/1970"){End_Date = "[No data]"}
    if(CPA_Margin == "0.00%" ){CPA_Margin = "[No data]"}
    if(Revenue == "0.00 kr" ){Revenue = "[No data]"}
    if(Predicted_Date_to_Goal == "01/01/1970" ){Predicted_Date_to_Goal = "[No data]"}
    if(_Client_Goal == "0.00%" ){_Client_Goal = "[No data]"}
    
    if (Delivery_Status == "Live" || Delivery_Status == "Paused"){
      var msg1 = prepareReport1(C_Link,Campaign_Name,Client,Delivery_Status,Start_Date,End_Date,Lead_Goal,Client_Goal_Type, Client_Delivery_Result, _Client_Goal, CPA_Margin, Predicted_Date_to_Goal, Revenue, Last_Received_Lead)
      allreportmsg.push(msg1)
      }
    
     if (Delivery_Status == "Live"){
      var msg2 = prepareReport2(Campaign_Name, Delivery_Status, End_Date, Lead_Goal, Client_Goal_Type,_Client_Goal,Avg_Client_Lead_per_day, Avg_Lead_Prediction)
      allreportmsg.push(msg2)
     }
    }
  allreportmsg.push("</body></html>")
  sendEmail(emailR, subjLine, allreportmsg.join(""))
}
