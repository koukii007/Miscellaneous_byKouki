# **Billboard Web-scraping Automation** 
## 1- Description:
This Automation scrapes the Chart of the Top 100 Artists on [Billboard.com](https://www.billboard.com/charts/artist-100/). It is Designed to get the 5 top artits of the Last 4 weeks including the current one.
## 2- Setup:
In Order to Run this Automation successfuly we would need to: 
1. Extract the wip file into an empty directory.
2. Open the directory with your text editor.
3. From the Terminal create a virtual environment : `python3 -m venv env`
4. cd to the newly created virtual environment and `pip` install the modules `scrapy` and `PyDrive`.

You can check if the web scrapper is working by simply running the python file `top100.py` from your text editor. This file should be able to be used without using `scrapy` through the PowerShell.
This script should generate `Artists.csv` after finishing scraping.

## 3- Getting Started with the Authentication on Google Cloud:
**This part is tricky so follow carefully!**

- Access the [Google Cloud Console](https://console.cloud.google.com/).
- Login With your Google Account.
- Create a new Project 

     ![alt text](1.png)
- Go to Enabled Apis

    ![alt text](/Pictures/2.png)

- Click on     ![alt text](/Pictures/3.png)
- In the Search Field look for **"drive api"**
- Choose the Google Drive API and **Enable** it.
- Now from the APIs and services drop down menu choose Credentials:
![alt text](/Pictures/4.png)

- Click : ![alt text](/Pictures/5.png) from the Top
 and choose this one ![alt text](/Pictures/6.png).

 - Now follow the OAuth consent screen and after completeting it will send you back to the same page which has the credentials.
 - You would have to click again on **Create Credentials**.
 - This time choose **Web Application** as Application Type and choose an appropriate name.
 - In the **Authorised JavaScript origins** type the following http://localhost:8080 :
 ![alt text](/Pictures/7.png)

 - In the **Authorised redirect URIs** Type the follwoing http://localhost:8080/ (front slash necessary):
 ![alt text](/Pictures/8.png)

 - Press **create** and a `.json` file should be genertated. Download it and move it to the directory where the scripts were extracted. 
 - Rename the `JSON` file to `client_secrets.json` .  **THIS IS IMPORTANT !**

 - Now you should be able to try to upload the `.csv` file created to your google drive. 

- run the `uploadToDrive.py` from your text editor and it should open up a new tab on your browser asking for an account to use to sign in.

- Leave it for now do no choose anything and go back to the console, to the **Credentials** page.
- In order for this to work you have to add yourself as User/Test.
- Click on the **OAuth counsent screen**  :
![alt text](/Pictures/10.png)
- Click on Add Users 
![alt text](/Pictures/11.png)
- Put your email and add your account.
- Now go back to the browser tab that was opened by the python script. Choose your account and sign in .
- To verify that you have successfuly connected your google drive you will get the following message:
![alt text](/Pictures/12.png)
- If you Check your google Drive you will find that the `.csv` file has been uploaded.
## 3 - Schedulling the Automation:
Now after testing the scripts and the setup its time to schedule this automation.

**On Windows you can use the Windows Task Scheduler**

But Assuming most servers run Linux, here is what to do:

1. Make Both `Python` Script files executable by adding `#!/usr/bin/python` to the first row.
2. On the terminal type:
- `chmod a+x top100.py`
- `chmod a+x uploadToDrive.py`
3. To check if you already have a crontab created use : `crontab -l`.
I fyou have never created a crontab this should show the message : `crontab: no crontab for ..`
4. Create a new Crontab by using : `crontab -e`
5. Now according to your needs use this [guide](https://crontab.guru/) and replace the asterisks depending on how often you would like this automation to run.
`* * * * * /path/top100.py >> ~/cron.log 2>&1`


## We are Almost done!
## 4- Form Generation And Data Sumamry:
In this Part we will go through how you can generate a Google Form from your csv File and Later how to Get the Responses from that form calculate the Average Ratings of Each Artist.

1. Go to [Google Drive](https://drive.google.com/drive/my-drive) and open the Google sheet.
2. Click on Extension and Choose **Apps Script**.
![alt text](/Pictures/13.png)
3. A new Tab will open a text editor with an empty function. 
4. Open the file From the Extracted Zip File Called `createForm.gs`

## createForm.gs
```javaScript
function myFunction() {
  
  const workingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const artists = workingSheet.getDataRange().getValues();

  Logger.log(artists);

  let form = FormApp.create('Top Artists Ranking');

  form.setTitle('Rate the Following Artists from 1 to 5')

  for (let i = 1; i < artists.length; i++) {
  form.addGridItem().setRows(artists[i]).setColumns(['1','2','3','4','5']);
  var items = form.getItems();
  var item = items[i-1];
  item.setTitle(artists[i])

}
  
  Logger.log('Published URL: ' + form.shortenFormUrl(form.getPublishedUrl()));
  workingSheet.getRange(1,4).setValue("Form Url:");
  workingSheet.getRange(1,5).setValue(form.getEditUrl());
  workingSheet.getRange(2,4).setValue("Form ID:");
  workingSheet.getRange(2,5).setValue(form.getId())
  }

```
5. `Copy Paste` the code from that file and replace the empty function with it.
6. Hit Run ![alt text](/Pictures/14.png) and this should write out a URL for the Generated Form that you can use to Share it with others.
- The Google Sheet will be Modified and the URL and ID will be added in Case needed in the Future.
7. Whenever you need to collect the Answers and calculate the Averages for each Rating , Click on the + sign and create a new `.gs` file ![alt text](/Pictures/15.png).
8. Open the provided script named `average.gs` and replace that empty Function with it.
## average.gs
```javaScript
function myFunction() {
  var workingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var formID = workingSheet.getRange(2,5).getValue();
  var myform = FormApp.openById(formID);
  var formResponses = myform.getResponses();
  if(formResponses.length == 0) return console.log('No responses') ;

  var scores = [];
  // Grab all Responses
  for(var i = 0; i < formResponses.length ; i++){
    var itemResponses = formResponses[i].getGradableItemResponses();
  
    for(var j = 0; j < itemResponses.length; j++){
      var itemResponse = itemResponses[j];
      var score = itemResponses[j].getResponse()[0];
  
      scores.push(itemResponse.getItem().getTitle());
      scores.push(score);
    
    }
  }
  Logger.log(scores)
  // Calculate Average
  var finalList = [];
  var sum = 0
  var count = 0
  for(var i = 0; i < scores.length; i+=2){
    sum = 0
    count =0
    var name = scores[i];
    for(var j=0; j < scores.length-1;j++){
      if(name == scores[j]){
        count+=1;
        sum += Number(scores[j+1]);
      } 
    }   
  if( !finalList.includes(name)){
      finalList.push(name);
      var average = sum/count
      finalList.push(average.toFixed(2));
    } 
  }
  Logger.log(finalList)
  //Write Back to the Sheet
  var column1 = workingSheet.getRange(1,1,10,1).getValues()
  workingSheet.getRange(1,2).setValue("Average Rating")
  for(var i=2;i < column1.length+1;i++){ 
    var data = workingSheet.getRange(i,1).getValue();
    for(var j=0; j <  finalList.length -1 ; j++){
      if(data == finalList[j]){
        workingSheet.getRange(i,2).setValue(finalList[j+1])
      }
    }

  }
}

```
9. Click Run and you will find the Averages Written in the fields that correspond to the Artist.

## You should be good to go now
For more information contact : koukii007.dev@gmail.com