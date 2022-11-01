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

The add-in will be reloaded by the webpack-dev-serveron each code save.

