# SheetsTA
This is a highly opinionated set of scripts and functions for Google Sheets, aimed at simplifying the process of grading student submissions. It is primarily created for courses in programming, where students submit links to their github repositories to assignments in Google Classroom.

## Installation
* Create a google apps script project, possibly contained in a Sheets spreadsheet
  * Make note of the project's id
* Download this project.
* Create a .clasp.json.prod in the project's base folder.
  * It should contain `{"scriptId":"your-id-here"}`
* Run `yarn install` to get all typings and modules needed to build the project
* Run `yarn run deploy:prod` to build & upload

Maybe this process will be simplified or better explained later. Don't hold your breath.

## Usage (not done)
* Basics
  * Getting list of active classrooms
  * Get stuff from a classroom
    * Roster
    * Assignments
  * Getting student submissions
  * Sanitizing Github links
* Document activity
  * Getting docs activity [dates/weeks]
  * Getting Github repo activity [dates/weeks]
* Master document
  * Setup/update everything
  * Updating roster
  * Updating submissions

## Plans
* Implementing the project as a proper standalone library?
* More granular control over how and if to filter activity dates for github links (using email adresses? Adding email adresses to the roster?) and docs pages
* Maybe some kind of optional subpages with readymade filters for github / docs submissions, and activity logging?
* Maybe some sort of simplified/visual way to generate _CONFIG page for master document?
* Finish this README / write some proper documentation