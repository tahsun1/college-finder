let collageData = [];
function showLoader() {
    document.querySelector(".loader-div").classList.add("shown");
}
function closeLoader() {
    document.querySelector(".loader-div").classList.remove("shown");
}
showLoader();
fetch("./data.xls")
    .then(function (response) {
        if (!response.ok) {
            throw new Error("Could not load data.xls");
        }

        return response.arrayBuffer();
    })

    .then(function (buffer) {
        var workbook = XLSX.read(buffer, {
            type: "array"
        });

        var sheet = workbook.Sheets[workbook.SheetNames[0]];

        var rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1
        });

        var keys = [
            "district",
            "thana",
            "eiin",
            "college_name",
            "shift",
            "version",
            "group",
            "gender",
            "min_GPA",
            "Own GPA",
            "SQ GPA",
            "available_seat"
        ];
        var result = rows.slice(1).map(function (row) {
            var obj = {};

            keys.forEach(function (key, index) {
                obj[key] = row[index] || "";
            });

            return obj;
        });

        collageData = result;
        collageData.splice(0, 1);
        console.log("Collages Loaded", collageData.length);
        displayCollage();
    })

    .catch(function (error) {
        console.error(error);

        document.getElementById("output").textContent =
            "Error: " + error.message;
    });

let collages = document.getElementById("collages");

let currentPage = 1;
const collegesPerPage = 50;
let currentCollageData = [];

function displayCollage() {
    currentCollageData = collageData;
    currentPage = 1;

    ((districts = []), (thanas = []));

    collageData.forEach(item => {
        districts.push(item.district);
        if (item.district == "") return;
        thanas.push(`${item.district}|${item.thana}`);
    });

    districts = [...new Set(districts)];
    thanas = [...new Set(thanas)];

    document.getElementById("district").innerHTML =
        "<option value=''>All</option>";

    districts.forEach(item => {
        let dst = document.createElement("option");
        dst.innerHTML = item;
        if (item == "") return;
        document.getElementById("district").appendChild(dst);
    });

    updateThana();

    displayPage();
}
function updateThana() {
    let dstValue = document.getElementById("district").value;
    document.getElementById("thana").innerHTML =
        "<option value=''>All</option>";
    let filteredThana = thanas.filter(function (i) {
        return !dstValue || i.split("|")[0] == dstValue;
    });

    filteredThana.forEach(function (item) {
        let tha = document.createElement("option");
        tha.innerHTML = item.split("|")[1];
        document.getElementById("thana").appendChild(tha);
    });
}
function displayFilteredCollage(filteredData) {
    currentCollageData = filteredData;
    currentPage = 1;
    displayPage();
}

function displayPage() {
    collages.innerHTML = "";

    let start = (currentPage - 1) * collegesPerPage;
    let end = Math.min(start + collegesPerPage, currentCollageData.length);
    for (let i = start; i < end; i++) {
        let item = currentCollageData[i];

        let tr = document.createElement("tr");

        tr.innerHTML = `
            <th>${i + 1}</th>
            <th>${item.district}</th>
            <th>${item.thana}</th>
            <th>${item.eiin}</th>
            <th>${item.college_name}</th>
            <th>${item.shift}</th>
            <th>${item.gender}</th>
            <th>${item.version}</th>
            <th>${item.group}</th>
            <th>${item.min_GPA}</th>
            <th>${item.available_seat}</th>
        `;

        collages.appendChild(tr);
    }
    
    document.querySelector('.note').innerHTML = `Showing ${collageData.length} Results`;
    displayPagination();

    closeLoader();
}

function displayPagination() {
    let totalPages = Math.ceil(currentCollageData.length / collegesPerPage);

    let footer = document.getElementById("pagination");

    footer.innerHTML = "";

    if (totalPages <= 1) return;

    // Previous
    let prev = document.createElement("button");
    prev.innerHTML = "&lt;";
    prev.disabled = currentPage === 1;
    prev.onclick = function () {
        currentPage--;
        displayPage();
    };
    footer.appendChild(prev);

    // Pages
    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, startPage + 4);

    if (startPage > 1) {
        addPageButton(1);

        if (startPage > 2) {
            addDots();
        }
    }

    for (let i = startPage; i <= endPage; i++) {
        addPageButton(i);
    }

    if (endPage < totalPages) {
        if (endPage < totalPages - 1) {
            addDots();
        }

        addPageButton(totalPages);
    }

    // Next
    let next = document.createElement("button");
    next.innerHTML = "&gt;";
    next.disabled = currentPage === totalPages;
    next.onclick = function () {
        currentPage++;
        displayPage();
    };
    footer.appendChild(next);

    function addPageButton(page) {
        let btn = document.createElement("button");

        btn.innerHTML = page;
        btn.disabled = page === currentPage;

        btn.onclick = function () {
            currentPage = page;
            displayPage();
        };

        footer.appendChild(btn);
    }

    function addDots() {
        let span = document.createElement("span");
        span.innerHTML = " ... ";
        footer.appendChild(span);
    }
}
function searchCollage(type) {
    let filteredName = document
        .getElementById("name")
        .value.trim()
        .toLowerCase();

    let gpaValue = document.getElementById("gpa").value.trim();

    let filteredGPA = gpaValue === "" ? null : parseFloat(gpaValue);

    let filteredEIIN = document.getElementById("eiinInput").value.trim();

    let filteredGroup = document
        .getElementById("group")
        .value.trim()
        .toLowerCase();

    let filteredDst = document
        .getElementById("district")
        .value.trim()
        .toLowerCase();

    let filteredThana = document
        .getElementById("thana")
        .value.trim()
        .toLowerCase();

    let filteredGender = document
        .getElementById("gender")
        .value.trim()
        .toLowerCase();

    let result = collageData.filter(function (college) {
        return (
            (!filteredName ||
                String(college.college_name)
                    .toLowerCase()
                    .startsWith(filteredName)) &&
            (filteredGPA === null || Number(college.min_GPA) === filteredGPA) &&
            (!filteredEIIN ||
                String(college.eiin).trim().startsWith(filteredEIIN)) &&
            (!filteredGroup ||
                String(college.group).toLowerCase().trim() === filteredGroup) &&
            (!filteredDst ||
                String(college.district).toLowerCase().trim() ===
                    filteredDst) &&
            (!filteredGender ||
                String(college.gender).toLowerCase().trim() ===
                    filteredGender) &&
            (!filteredThana ||
                String(college.thana).toLowerCase().trim() === filteredThana)
        );
    });
    displayFilteredCollage(result);
}
function scrollToTop() {
    document.querySelector(".form").scrollTo({
        top: 0,
        behavior: "smooth"
    });
}
function displayFilteredCollage(filteredData) {
    currentCollageData = filteredData;
    currentPage = 1;

    function displayPage() {
        collages.innerHTML = "";

        let start = (currentPage - 1) * collegesPerPage;
        let end = Math.min(start + collegesPerPage, filteredData.length);

        for (let i = start; i < end; i++) {
            let item = filteredData[i];

            let tr = document.createElement("tr");

            tr.innerHTML = `
                <th>${i + 1}</th>
                <th>${item.district}</th>
                <th>${item.thana}</th>
                <th>${item.eiin}</th>
                <th>${item.college_name}</th>
                <th>${item.shift}</th>
                <th>${item.gender}</th>
                <th>${item.version}</th>
                <th>${item.group}</th>
                <th>${item.min_GPA}</th>
                <th>${item.available_seat}</th>
            `;

            collages.appendChild(tr);
        }
        
        let resultCount = filteredData.length;
        document.querySelector('.note').innerHTML = `Showing ${resultCount} Results`;
        displayPagination();
    }

    function displayPagination() {
        let totalPages = Math.ceil(filteredData.length / collegesPerPage);

        let footer = document.getElementById("pagination");
        footer.innerHTML = "";

        if (totalPages <= 1) return;

        let prev = document.createElement("button");
        prev.innerHTML = "&lt;";
        prev.disabled = currentPage === 1;

        prev.onclick = function () {
            currentPage--;
            displayPage();
        };

        footer.appendChild(prev);

        let startPage = Math.max(1, currentPage - 2);
        let endPage = Math.min(totalPages, startPage + 4);

        if (startPage > 1) {
            addPageButton(1);

            if (startPage > 2) {
                addDots();
            }
        }

        for (let i = startPage; i <= endPage; i++) {
            addPageButton(i);
        }

        if (endPage < totalPages) {
            if (endPage < totalPages - 1) {
                addDots();
            }

            addPageButton(totalPages);
        }

        let next = document.createElement("button");
        next.innerHTML = "&gt;";
        next.disabled = currentPage === totalPages;

        next.onclick = function () {
            currentPage++;
            displayPage();
        };

        footer.appendChild(next);

        function addPageButton(page) {
            let btn = document.createElement("button");

            btn.innerHTML = page;
            btn.disabled = page === currentPage;

            btn.onclick = function () {
                currentPage = page;
                displayPage();
            };

            footer.appendChild(btn);
        }

        function addDots() {
            let span = document.createElement("span");
            span.innerHTML = " ... ";
            footer.appendChild(span);
        }
    }

    displayPage();
}
function resetForm() {
    document.getElementById("district").value = "";
    document.getElementById("gpa").value = "";
    document.getElementById("eiinInput").value = "";
    document.getElementById("name").value = "";
    document.getElementById("group").value = "";
    document.getElementById("gender").value = "";
    updateThana();
    displayCollage();
}
