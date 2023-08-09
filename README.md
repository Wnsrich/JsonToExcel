# PCF_JsonToExcel

Details:

![image](https://github.com/Wnsrich/PCF_JsonToExcel/assets/103210745/e6849bfe-1d63-4969-9e86-2285e30fcdfe)



Trigger Method: 

  use a variable as trigger, when the trigger value eq true, the data will transfer to excel:

Sample in Canvas App:

  Set(JSONContent,JSON(colData));
  
  Set(sortOrder,Text("[""name"",""city"",""years""]"));
  
  UpdateContext({trigger:false});
  
  UpdateContext({trigger:true});
  

Results:

![image](https://github.com/Wnsrich/PCF_JsonToExcel/assets/103210745/6bcc0945-1501-4740-9e5b-cb6ad1118a63)



Improve:
 - excel column width auto-adjust
 - button trigger
 - fast response
