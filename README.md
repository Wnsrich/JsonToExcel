# PCF_JsonToExcel

Details:

![image](https://github.com/Wnsrich/PCF_JsonToExcel/assets/103210745/e13acfa5-0155-4513-af66-cccfa491f55f)

Trigger Method: 
  use a variable as trigger, when the trigger value eq true, the data will transfer to excel:

Sample in Canvas App:

  Set(JSONContent,JSON(colData));
  Set(sortOrder,Text("[""name"",""city"",""years""]"));
  UpdateContext({trigger:false});
  UpdateContext({trigger:true});

Results:

![image](https://github.com/Wnsrich/PCF_JsonToExcel/assets/103210745/ab147fd0-a983-485e-aaf0-383c9b53a5bc)

Improve:
 - excel column width auto-adjust
 - button trigger
 - fast response
