let courseData, courses, gradeData, grades;
async function loadPage() {
    try {
        courseData = await getExcelFile("./AU_course_credit.xlsx");
        courses = getColumnData(courseData, "COURSE TITLE");

        gradeData = await getExcelFile("./grade_point_conversion.xlsx");
        grades = getColumnData(gradeData, "Letter Grade");

        showCourses(courses, "course");
    } catch (error) {
        console.error("Error loading Excel files: " + error);
    }
}

async function getExcelFile(path) {
    try {
        const response = await fetch(path);
        const data = await response.arrayBuffer();
        const uint8Array = new Uint8Array(data);
        const workbook = XLSX.read(uint8Array, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        return jsonData;
    } catch (error) {
        console.error("Error reading file " + error);
        return [];
    }
}

function getColumnData(data, column) {
    return data.map(item => item[column]);
}

function showCourses(courses, id) {
    if (Array.isArray(courses)) {
        const selectElement = document.getElementById(id);
        selectElement.innerHTML = "";

        courses.forEach(course => {
            const option = document.createElement('option');
            option.value = course;
            option.textContent = course;
            selectElement.appendChild(option);
        });

        const excelDataElement = document.getElementById("excelData");
        excelDataElement.innerText = JSON.stringify(courses, null, 2);
    } else {
        console.error("Invalid data format: 'courses' is not an array.");
    }
}

function addCourse() {
    const selectedCourses = Array.from(document.getElementById('course').selectedOptions);
    const courseTable = document.getElementById('courseTable');

    // Clear existing rows in the courseTable
    courseTable.innerHTML = "";

    selectedCourses.forEach(selectedCourse => {
        const newRow = document.createElement("tr");

        const courseColumn = document.createElement("td");
        courseColumn.textContent = selectedCourse.value;
        newRow.appendChild(courseColumn);

        const gradeColumn = document.createElement("td");
        const gradeSelect = document.createElement("select");

        grades.forEach(grade => {
            const option = document.createElement("option");
            option.value = grade;
            option.textContent = grade;
            gradeSelect.appendChild(option);
        });

        gradeColumn.appendChild(gradeSelect);
        newRow.appendChild(gradeColumn);

        const actionColumn = document.createElement("td");
        const deleteButton = document.createElement("button");
        deleteButton.textContent = "Delete";

        deleteButton.addEventListener("click", function() {
            courseTable.removeChild(newRow);
        });

        actionColumn.appendChild(deleteButton);
        newRow.appendChild(actionColumn);

        courseTable.appendChild(newRow);
    });
}


function convertData(data, detail) {
    const dataToConvert = detail === 'course' ? courseData : gradeData;
    const dataKey = detail === 'course' ? 'COURSE TITLE' : 'Letter Grade';
    const matchingItem = dataToConvert.find(item => item[dataKey] === data);

    return matchingItem[detail === 'course' ? 'CREDITS' : 'Grade Point'];
}

function getTableData() {
    const tableData = [];
    const rows = document.querySelectorAll('#gradeTable tbody tr');

    rows.forEach(row => {
        const [course, grade] = row.querySelectorAll('td');
        tableData.push({
            creditPoint: convertData(course.textContent, 'course'),
            gradePoint: convertData(grade.querySelector('select').value, 'grade')
        });
    });

    return tableData;
}



function cgpa() {
    var num = 0;
    var den = 0;
    const txt = "Your CGPA is ";
    const data = getTableData();

    data.forEach(entry => {
        const creditPoint = entry.creditPoint;
        const gradePoint = entry.gradePoint;

        if (gradePoint !== 0) {
            num += creditPoint * gradePoint;
            den += creditPoint;
        }
    });

    const cgpaValue = den === 0 ? 0 : num / den;
    document.getElementById('cgpa').innerHTML = txt + cgpaValue.toFixed(2);
    console.log(data);
}