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

The add-in is deployed on [netlify](https://office-add-in-2-7d7ee5.netlify.app).

Quick test in a browser: [click](https://office-add-in-2-7d7ee5.netlify.app/taskpane.html) to open the add-in's Welcome page.

Test in the MS Word web app:

- in Word, Insert -> Add-ins -> Upload my Add-in, navigate to a local copy of the [manifest-netlify](https://github.com/rudifa/Office-Add-in-2/blob/main/manifest-netlify.xml) and upload it
- you may need one or two web page restarts until the add-in icon appears as the rightmost item in the Ribbon
- try it.

### Testing with the netlify and the MS Word desktop

Not yet done at this stage.

The procedure should be similar to the above procedure for the web.

### References

[Tutorial: Create a Word task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial)

[Word package js API reference](https://learn.microsoft.com/en-us/javascript/api/word?view=word-js-preview)

[Office Add-In - Xomino](https://xomino.com/category/office-add-in/) - tables, ... 
