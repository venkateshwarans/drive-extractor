# drive-extractor

### Installation Steps
**1.** Install `clasp` globally 
```
npm install @google/clasp -g
clasp login // to login to your Google account for authentication
```

**2.** Clone the repo 
```
git clone https://github.com/venkateshwarans/drive-extractor.git && cd drive-extractor
npm install
```

### How to run the addon
**1.** Open a new spreadsheet - (click here [sheet.new](https://sheet.new/))

**2.** Under 'Tools', click 'Script Editor'. It will create an Untitle Project. 

**3** Click 'Project Settings' from the left menu, and copy the `Script ID`

**4.** Open `.clasp.json` and paste the above `scriptId`. 


### How to build and upload to Google Apps Script
**1.**```
npm run build```


**2.**```
npm run upload```

### How to execute 
**1.** Go back to scripts.google.com

**2.** Open the project that you just created.

**3.** Hit 'Run'. It will ask for your Google Login. Once you give permission, it will execute the script and you will see the Add on under the 'Add-ons' menu in the spreadsheet
