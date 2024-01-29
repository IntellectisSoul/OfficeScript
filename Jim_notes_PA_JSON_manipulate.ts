
1. Qn : i have a JSON object in memory like this : 
{
    "firstname": "Alex",
    "age": 34,
    "active": true,
    "location": null,
    "job": {
      "role": "technical support engineer",
      "employer": "Microsoft",
      "yearsActive": [
        2017,
        2018,
        2022,
        2023
      ]
    }
  }

  i wish to change the value of the property firstname to 'Jim'

  {  "firstname": "Jim"
    "age": 34,
    "active": true,
    "location": null,
    "job": {
      "role": "technical support engineer",
      "employer": "Microsoft",
      "yearsActive": [
        2017,
        2018,
        2022,
        2023
      ]
    },
    }

 
Answer : correct JSON manipulation action : 
setProperty(variables('jsonObject'),'firstname','Jim')

Note : the order may be changed but this is OK and does not affect the object's integrity.

  
-----


2. Question : Expression to add Properties to extg. Object in above example : 
a.Initialize variable 'order' as : 
[
    "firstname",
    "age",
    "active",
    "location",
    "job"
  ]
b.Initialize variable 'data' as original object  
c. Initialize finalStructure variable type as Object as follows : 
    json(concat('{"order":', string(variables('order')) , ',"data":', string(variables('data')), '}'))


Output : 
{"order":["firstname","age","active","location","job"],
"data":{
    "firstname":"Alex",
    "age":34,
    "active":true,
    "location":null,
    "job":{"role":"technical support engineer",
    "employer":"Microsoft",
    "yearsActive":[2017,2018,2022,2023]
}
}
}

2.1 Question : to manipulate and change the value of a property firstname from Alex to Joseph depicted as a variable named firtname1 : 

Answer primitive : setProperty(variables('jsonObject'), 'firstname', variables('firstname1'))
    Note : the property inside the Object is now referred to as variables().

{
    "order": [
      "firstname",
      "age",
      "active",
      "location",
      "job"
    ],
    "data": {
      "age": 34,
      "active": true,
      "location": null,
      "job": {
        "role": "technical support engineer",
        "employer": "Microsoft",
        "yearsActive": [
          2017,
          2018,
          2022,
          2023
        ]
      },
      "firstname": "Joseph"
    }
  }


----
3. Looping through the array (from List rows present in a table, saved to an array variable)

 a. Compose : setProperty(item(), 'Count of Concur', outputs('Concur_Count'))
 b. Append to array variable : outputs('Update_arr_MasterData_ConcurCount')  : you need to save to a variable in order to persist or you will lose data after the loop ends.


---


  4.when you need information about the iteration itself, like the current index.
  iterationIndexes('Your_Apply_to_each_Loop')['currentIndex']
--


setProperty(variables('obj_currentRow'), 'Approver GEID', items('Apply_to_each_4')['Approver GEID'])




4. Below is the output from Officescript 'Jim_find_startingRange_createTable.ts', which returns {tableValuesString };
{
    "tableValuesString": "[[\"Employee ID\",\"Report Name\",\"Submitted Date\",\"Approval Status\",\"Approver Employee ID\",\"Approver Email Address\",\"Amount\"],[\"1200248\",\"MyClaim-Mar2021\",44317,\"Pending CCM Approval\",1200060,\"jowthean@ncs.com.sg\",150],[\"1205953\",\"Transpo Claim - MarApr2021\",44317,\"Pending CCM Approval\",\"1200054\",\"lyndyng@ncs.com.sg\",58.4],[\"1315481\",\"PNM1-P1315481-202104-Transpo\",44317,\"Pending CCM Approval\",\"1200076\",\"henrylu@ncs.com.sg\",92.44],[\"1325359\",\"IRAS Claims\",44317,\"Pending CCM Approval\",\"1200150\",\"vanessat@ncs.com.sg\",86.04]]"
  }
after conversion using : json(outputs('Run_script_from_SharePoint_library')?['body/result']['tableValuesString']), i get the structure below : 
[
  [
    "Employee ID",
    "Report Name",
    "Submitted Date",
    "Approval Status",
    "Approver Employee ID",
    "Approver Email Address",
    "Amount"
  ],
  [
    "1200248",
    "MyClaim-Mar2021",
    44317,
    "Pending CCM Approval",
    1200060,
    "jowthean@ncs.com.sg",
    150
  ],
  [
    "1205953",
    "Transpo Claim - MarApr2021",
    44317,
    "Pending CCM Approval",
    "1200054",
    "lyndyng@ncs.com.sg",
    58.4
  ],
  [
    "1315481",
    "PNM1-P1315481-202104-Transpo",
    44317,
    "Pending CCM Approval",
    "1200076",
    "henrylu@ncs.com.sg",
    92.44
  ],
  [
    "1325359",
    "IRAS Claims",
    44317,
    "Pending CCM Approval",
    "1200150",
    "vanessat@ncs.com.sg",
    86.04
  ]
]