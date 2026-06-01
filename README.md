# DSRUEP - Data Science and Remote Understanding Educational Platform

A comprehensive educational platform combining interactive lesson websites with Google Classroom automation tools. This repository contains both interactive STEM lessons and administrative scripts to streamline classroom management and curriculum development.

## 📚 Contents

### Interactive Lessons
Located in the `/m5` directory, these are self-contained, single-file HTML lesson websites designed for easy deployment and student engagement.

- **quarks.html** - An interactive physics lesson on "Quarks: The Building Blocks of Matter" featuring:
  - Russian Doll analogy for understanding nested particle structures
  - Interactive particle builder for hands-on learning
  - Quiz and assessment components
  - Styled for ESL and mixed-ability students

### Administrative Tools
Located in the `/Admin-26` directory, these Google Apps Script utilities automate classroom management and curriculum development tasks.

#### Classroom Management Scripts
- **Classroom-Assignment-Automation.js** - Automates creation of Google Classroom assignments from a spreadsheet with fields for post date, title, description, course ID, due date, points, and materials
- **attendance-script.js** - Generates attendance sheets for a date range, automatically skipping weekends
- **classroom-grades-26-1.js** - Manages and processes classroom grade data
- **pull-class-codes.js** - Retrieves class codes from Google Classroom

#### Curriculum Development Tools
Located in `/Admin-26/lesson-creator-m3`:
- **parse-pdf-to-json.js** - Extracts text from PDF files using Google Drive OCR conversion
- **parse-pdf-to-json-and-generate-lessons.js** - Combines PDF parsing with lesson generation
- **CurriculumGenerator.js** - Generates complete curriculum structures

## 🚀 Features

- **Interactive Learning**: Single-file HTML lessons with Tailwind CSS styling for responsive design
- **Hands-on Activities**: Drag-and-drop interactions and interactive quiz components
- **Classroom Integration**: Seamless integration with Google Classroom for assignment and grade management
- **Automation**: Scripts to automate routine classroom tasks and save instructor time
- **Curriculum Tools**: PDF-to-lesson conversion utilities for rapid curriculum development

## 📝 License

See LICENSE file for details.


