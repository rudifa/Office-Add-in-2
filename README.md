# Learning the development of MS Office Add-ins

This project was created and developed following the MS tutorial [Create a Word task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial).

### Testing with the localhost and the MS Word - web

- open the MS word [in the browser](https://www.office.com/)
- log in if needed
- click on Word icon
- open a document
- from Share button copy the link similar to `https://1drv.ms/w/s!AoxMqCIhJWzTgQkVbzUJ1f2DAWDG?e=dAbERv`
- in terminal, run `npm run start:web -- --document https://1drv.ms/w/s!AoxMqCIhJWzTgQkVbzUJ1f2DAWDG?e=dAbERv`
- in Word, accept the installation of the add-in
- click on the  `Show Taskpane` icon

Normally, the webpack-dev-server reloads the add-in after each code change.
If in doubt:

-  force reload the MS Word page in the browser
-  stop and restart the server:

 ```
 mpm stop
 npm run start:...
 ``` 

### Testing with the netlify and the MS Word - web

The add-in is deployed on [netlify](https://office-add-in-2-7d7ee5.netlify.app/taskpane.html).

Test in the MS Word web app:

1. download a copy  of the file [manifest-netlify.xml](https://office-add-in-2-7d7ee5.netlify.app/assets/manifest-netlify.xml) from my github repo.
2. open [MS Office](https://www.office.com/)
3. log in or sign up
4. open the Word app
5. open a new document
6. in `Insert -> Add-ins -> Upload my Add-in` navigate to the local copy of the `manifest-netlify.xml` and upload it
7. a `Show Taskpane` icon will appear at the right of the Ribbon
8. click the icon to open the Task Pane
9. try it


### Testing with the netlify and the MS Word desktop on Mac

Follow the [instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac), briefly:

1. download the file [manifest-netlify.xml](https://office-add-in-2-7d7ee5.netlify.app/assets/manifest-netlify) and copy it into the designated [directory](/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef)
2. in the desktop Word app, go to the Insert - Add-ins and install the add-in
3. try it

### References

[Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial)

[Word package js API reference](https://learn.microsoft.com/en-us/javascript/api/word?view=word-js-preview)

[Office Add-In - Xomino](https://xomino.com/category/office-add-in/) - tables, ... 
