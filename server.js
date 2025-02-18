const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;

async function getStudentData(rollNo) {
    try {
        // Add headers to mimic browser request
        const headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        };

        const studentResponse = await axios.post(
            `https://results.bput.ac.in/student-detsils-results?rollNo=${rollNo}`,
            { headers }
        );

        const gradesResponse = await axios.post(
            `https://results.bput.ac.in/student-results-subjects-list?semid=2&rollNo=${rollNo}&session=Even-(2021-22)`,
            { headers }
        );

        // Check if the responses contain actual data
        if (!studentResponse.data || !gradesResponse.data) {
            throw new Error('No data received from API');
        }

        return {
            name: studentResponse.data.studentName,
            rollNo: rollNo,
            subjects: gradesResponse.data.map(subject => ({
                subjectName: subject.subjectName,
                grade: subject.grade
            }))
        };
    } catch (error) {
        console.error(`Error fetching data for roll no ${rollNo}:`, error.message);
        return null;
    }
}

app.get('/generate-grades', async (req, res) => {
    try {
        const students = [];
        const subjectSet = new Set();

        // Add delay between requests to prevent rate limiting
        const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

        // Fetch data for roll numbers 22 to 81
        for (let i = 22; i <= 81; i++) {
            const rollNo = `21011100${i.toString().padStart(2, '0')}`;
            console.log(`Fetching data for ${rollNo}...`);
            
            const studentData = await getStudentData(rollNo);
            await delay(1000); // 1 second delay between requests
            
            if (studentData && studentData.subjects.length > 0) {
                students.push(studentData);
                studentData.subjects.forEach(subject => {
                    subjectSet.add(subject.subjectName);
                });
            }
        }

        // Create Excel data with roll numbers
        const subjects = Array.from(subjectSet);
        const excelData = students.map(student => {
            const row = {
                'Roll Number': student.rollNo,
                'Student Name': student.name
            };
            
            subjects.forEach(subject => {
                row[subject] = '';
            });

            student.subjects.forEach(subjectData => {
                row[subjectData.subjectName] = subjectData.grade;
            });

            return row;
        });

        // Create and save Excel file
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(excelData);
        
        // Set column widths
        const columnWidths = {
            'A': 15, // Roll Number
            'B': 30  // Student Name
        };
        worksheet['!cols'] = Object.values(columnWidths);

        xlsx.utils.book_append_sheet(workbook, worksheet, "Student Grades");
        
        const fileName = "studentGrades.xlsx";
        const filePath = path.join(__dirname, fileName);
        xlsx.writeFile(workbook, filePath);

        res.json({ 
            message: "Grade sheet generated successfully!",
            filePath: filePath,
            studentsProcessed: students.length
        });

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: "Failed to generate grade sheet" });
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});