const dayMap = { "A": "Saturday", "S": "Sunday", "M": "Monday", "T": "Tuesday", "W": "Wednesday", "R": "Thursday", "F": "Friday" };
let extractedCourses = [];
let facultyData = {};
let uniqueTimeSlots = new Set();
let foundDays = new Set();

document.getElementById('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
        const ab = await file.arrayBuffer();
        const wb = XLSX.read(ab, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        processExcelData(data);
    } catch (err) {
        alert("Error reading file.");
    }
});

function processExcelData(rows) {
    extractedCourses = [];
    uniqueTimeSlots.clear();
    foundDays.clear();
    const coursesFound = new Set();

    let currentFullCode = "";
    let currentBaseCode = "";
    let currentSection = "-";

    rows.forEach((row) => {
        let rowStr = row.map(c => c ? c.toString().trim() : "").join(" ");

        if (rowStr.includes("Name:")) {
            let nameMatch = rowStr.match(/Name:\s*(.*?)\s*ID#/i);
            if (nameMatch) document.getElementById('disp-name').innerText = "Name: " + nameMatch[1].trim();
            let idMatch = rowStr.match(/ID#\s*([\d-]+)/i);
            if (idMatch) document.getElementById('disp-id').innerText = "ID#: " + idMatch[1].trim();
        }
        const semMatch = rowStr.match(/(Spring|Summer|Fall)\s*-\s*\d{4}/i);
        if (semMatch) document.getElementById('disp-semester').innerText = "Semester: " + semMatch[0].toUpperCase();

        const courseMatch = rowStr.match(/([A-Z]{2,4}\d{3,4}(?:\sLab)?)/i);
        const timeMatch = rowStr.match(/([A-Z]+)\s+(\d{1,2}:\d{2}[APM]+\s?-\s?\d{1,2}:\d{2}[APM]+)/i);

        if (courseMatch) {
            currentFullCode = courseMatch[1].trim();
            currentBaseCode = currentFullCode.replace(/\sLab/i, "").trim();
            
            let courseIdx = row.findIndex(c => c && c.toString().includes(currentFullCode));
            if (courseIdx !== -1) {
                for (let i = courseIdx + 1; i < row.length; i++) {
                    let cellVal = row[i] ? row[i].toString().trim() : "";
                    if (cellVal && /^\d+$/.test(cellVal)) { 
                        currentSection = cellVal; 
                        break; 
                    }
                }
            }
            coursesFound.add(currentBaseCode);
        }

        if (currentFullCode && timeMatch) {
            const days = timeMatch[1].toUpperCase();
            const time = timeMatch[2].toUpperCase().replace(/\s/g, '');
            const roomMatch = rowStr.match(/\d{3}\s?\(.*?\)|AB\d-\d{3}|FUB-\d{3}/i);
            const room = roomMatch ? roomMatch[0] : "TBA";

            uniqueTimeSlots.add(time);

            for (let char of days) {
                if (dayMap[char]) {
                    foundDays.add(dayMap[char]);
                    extractedCourses.push({ 
                        code: currentFullCode, 
                        baseCode: currentBaseCode,
                        section: currentSection, 
                        time: time, 
                        room: room, 
                        activeDay: dayMap[char] 
                    });
                }
            }
        }
    });

    renderTable();
    if (coursesFound.size > 0) showFacultyModal(Array.from(coursesFound));
}

function parseTime(t) {
    let timeStr = t.split('-')[0]; 
    let modifier = timeStr.slice(-2);
    let time = timeStr.slice(0, -2);
    let [hours, minutes] = time.split(':');
    if (hours === '12') hours = (modifier === 'AM') ? '00' : '12';
    else if (modifier === 'PM') hours = parseInt(hours, 10) + 12;
    return `${hours.toString().padStart(2, '0')}:${minutes}`;
}

function renderTable() {
    const sortedTimes = Array.from(uniqueTimeSlots).sort((a, b) => parseTime(a).localeCompare(parseTime(b)));
    const head = document.getElementById('time-header');
    head.innerHTML = `<th class="bg-blue-900 border-r border-blue-800 text-2xl p-6">Day / Time</th>`;
    sortedTimes.forEach(t => head.innerHTML += `<th class="text-sm font-black p-5 uppercase tracking-tighter">${t}</th>`);

    const body = document.getElementById('routineBody');
    body.innerHTML = "";
    const daysOrder = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].filter(d => foundDays.has(d));

    daysOrder.forEach(day => {
        let row = `<tr><td class="bg-slate-100 dark:bg-slate-800 font-black text-blue-900 dark:text-blue-400 uppercase text-2xl p-6 border-r">${day}</td>`;
        sortedTimes.forEach(time => {
            const matches = extractedCourses.filter(c => c.activeDay === day && c.time === time);
            if (matches.length > 0) {
                let cellHtml = `<td class="bg-blue-50/30 dark:bg-blue-900/10 p-2">`;
                matches.forEach(m => {
                    const fac = facultyData[m.baseCode] ? `<div class="mt-2 bg-blue-600 text-white px-2 py-1 rounded text-[10px] font-bold block w-fit mx-auto">${facultyData[m.baseCode]}</div>` : "";
                    cellHtml += `
                        <div class="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-blue-100 dark:border-slate-700 mb-2">
                            <div class="text-blue-900 dark:text-blue-300 font-black text-lg">${m.code}</div>
                            ${fac}
                            <div class="text-[12px] text-slate-600 dark:text-slate-400 font-bold mt-2 uppercase">SEC: ${m.section} | RM: ${m.room}</div>
                        </div>`;
                });
                cellHtml += `</td>`;
                row += cellHtml;
            } else { row += `<td class="text-slate-300 dark:text-slate-700 font-bold">-</td>`; }
        });
        row += `</tr>`;
        body.innerHTML += row;
    });
    document.getElementById('upload-box').classList.add('hidden');
    document.getElementById('routineContainer').classList.remove('hidden');
}

function openEditModal() {
    const list = document.getElementById('editCourseList');
    list.innerHTML = "";
    extractedCourses.forEach((c, index) => {
        list.innerHTML += createCourseRow(c, index);
    });
    document.getElementById('editModal').classList.remove('hidden');
}

function createCourseRow(c, index) {
    return `
    <div class="bg-slate-50 dark:bg-slate-800 p-4 rounded-2xl border border-slate-200 dark:border-slate-700 flex flex-wrap gap-3 items-center">
        <div class="flex-grow grid grid-cols-2 md:grid-cols-5 gap-3">
            <input type="text" value="${c.code}" placeholder="Course" class="ed-code p-2 rounded border dark:bg-slate-900 text-sm font-bold">
            <input type="text" value="${c.section}" placeholder="Section" class="ed-sec p-2 rounded border dark:bg-slate-900 text-sm">
            <input type="text" value="${c.room}" placeholder="Room" class="ed-room p-2 rounded border dark:bg-slate-900 text-sm">
            <select class="ed-day p-2 rounded border dark:bg-slate-900 text-sm">
                ${Object.values(dayMap).map(d => `<option value="${d}" ${d===c.activeDay?'selected':''}>${d}</option>`).join('')}
            </select>
            <input type="text" value="${c.time}" placeholder="8:00AM-10:30AM" class="ed-time p-2 rounded border dark:bg-slate-900 text-sm">
        </div>
        <button onclick="this.parentElement.remove()" class="bg-red-100 text-red-600 p-2 rounded-lg hover:bg-red-200">Remove</button>
    </div>`;
}

function addNewCourseRow() {
    const list = document.getElementById('editCourseList');
    list.insertAdjacentHTML('beforeend', createCourseRow({code:'', section:'', room:'', activeDay:'Saturday', time:''}));
}

function saveEdits() {
    const rows = document.querySelectorAll('#editCourseList > div');
    extractedCourses = [];
    uniqueTimeSlots.clear();
    foundDays.clear();

    rows.forEach(row => {
        const code = row.querySelector('.ed-code').value.toUpperCase();
        const time = row.querySelector('.ed-time').value.toUpperCase().replace(/\s/g, '');
        const day = row.querySelector('.ed-day').value;
        if (code && time) {
            uniqueTimeSlots.add(time);
            foundDays.add(day);
            extractedCourses.push({
                code: code,
                baseCode: code.replace(/\sLab/i, "").trim(),
                section: row.querySelector('.ed-sec').value,
                room: row.querySelector('.ed-room').value.toUpperCase(),
                time: time,
                activeDay: day
            });
        }
    });
    renderTable();
    document.getElementById('editModal').classList.add('hidden');
}

function showFacultyModal(courseList) {
    const container = document.getElementById('facultyInputs');
    container.innerHTML = "";
    courseList.forEach(c => {
        container.innerHTML += `<div class="bg-slate-50 dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700"><label class="block text-xs font-black text-blue-600 mb-1 uppercase">${c}</label><input type="text" placeholder="Faculty Initial" class="fac-in w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 p-2 rounded-lg text-sm" data-course="${c}"></div>`;
    });
    document.getElementById('facultyModal').classList.remove('hidden');
}

function applyFaculty(isSkip = false) {
    facultyData = {};
    if (!isSkip) {
        document.querySelectorAll('.fac-in').forEach(i => {
            if (i.value) facultyData[i.dataset.course] = i.value.toUpperCase();
        });
    }
    document.getElementById('facultyModal').classList.add('hidden');
    renderTable();
}

function downloadAsImage() {
    const routine = document.getElementById('routineContainer');
    const htmlTag = document.documentElement;
    
    // ১. চেক করা হচ্ছে বর্তমানে ডার্ক মোড অন আছে কি না
    const isDarkMode = htmlTag.classList.contains('dark');

    // ২. ডাউনলোড করার সময় সাময়িকভাবে ডার্ক মোড অফ করা (সাদা করার জন্য)
    if (isDarkMode) {
        htmlTag.classList.remove('dark');
    }

    // ৩. মোবাইল ভিউ ঠিক রাখার জন্য উইথ সেট করা
    const originalStyle = routine.style.cssText;
    routine.style.width = "1200px"; 
    routine.style.minWidth = "1200px";

    // ৪. ছবি ক্যাপচার করা
    html2canvas(routine, { 
        scale: 3, 
        backgroundColor: "#ffffff", // ব্যাকগ্রাউন্ড সাদা নিশ্চিত করা
        windowWidth: 1200,
        ignoreElements: (el) => el.classList.contains('no-print') 
    }).then(canvas => {
        // ৫. ছবি তোলা শেষ হলে ডার্ক মোড আবার ফিরিয়ে আনা (যদি আগে অন ছিল)
        if (isDarkMode) {
            htmlTag.classList.add('dark');
        }

        // ৬. স্টাইল আগের মতো করে দেওয়া
        routine.style.cssText = originalStyle;

        const a = document.createElement('a');
        a.download = 'EWU_Smart_Routine.png';
        a.href = canvas.toDataURL("image/png");
        a.click();
    });
}

document.addEventListener('contextmenu', event => event.preventDefault());

document.onkeydown = function(e) {
    if (e.keyCode == 123) return false;
    if (e.ctrlKey && e.shiftKey && e.keyCode == 'I'.charCodeAt(0)) return false;
    if (e.ctrlKey && e.shiftKey && e.keyCode == 'J'.charCodeAt(0)) return false;
    if (e.ctrlKey && e.keyCode == 'U'.charCodeAt(0)) return false;
    if (e.ctrlKey && e.keyCode == 'S'.charCodeAt(0)) return false;
    if (e.ctrlKey && e.keyCode == 'C'.charCodeAt(0)) {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') return false;
    }
};