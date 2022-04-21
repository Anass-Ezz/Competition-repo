var type_input = document.getElementById("type");
var title_input = document.getElementById("title");
var desc_input = document.getElementById("desc");
var date_input = document.getElementById("date");
var bg_input = document.getElementById("bg");
var excel_input = document.getElementById("excel");

var primary_color = document.getElementById("primary-color").value;
var secondary_color = document.getElementById("secondary-color").value;


var alert_element = document.getElementById("alert");
var preview_button = document.getElementById("preview");
var student_table = document.getElementById("results");

let fablab_url 
let greenlab_url 
fetch("images.json")
    .then(response => response.json())
    .then(json => {fablab_url = json.fablab; greenlab_url = json.greenlab});


function show_message(message, type) {
    let curret_type
    alert_element.textContent = message;
    alert_element.classList = ''
    alert_element.classList.add('alert', type)
    curret_type = type
    console.log(curret_type)
    alert_element.setAttribute("reveal", "");
    window.setTimeout(() => {
        alert_element.removeAttribute("reveal");
    }, 4000);
}
function test_fields() {
    if (
        type_input.value &&
        title_input.value &&
        desc_input.value &&
        date_input.value &&
        bg_input.value &&
        excel_input.value
    ) {
        return true;
    } else {
        show_message("Fill all the fields please", "alert-danger");
        return false;
    }
}

var table = document.getElementById("table_body");
var data;
function generate_table() {

    if (test_fields()) {
        student_table.hidden = false
        readXlsxFile(excel_input.files[0]).then((data) => {
            data = data;
            data.forEach((student) => {
                console.log(student);
                var table_row = document.createElement("tr");
                var name_element = document.createElement("td");
                name_element.textContent = student[0];
                var last_name_element = document.createElement("td");
                last_name_element.textContent = student[1];
                var row_actions = document.createElement("td");
                var download_button = document.createElement("button");
                download_button.textContent = "Download";
                download_button.classList.add("btn", "btn-primary");
                download_button.addEventListener("click", () => {
                    generatePDF_download(student[0], student[1]);
                });
                var send_button = document.createElement("button");
                send_button.textContent = "Send Email";
                send_button.classList.add("btn", "btn-danger");
                send_button.addEventListener("click", () => {
                    send_solo_email(student[0], student[1], student[2]);
                });
                row_actions.classList.add("row-actions");
                row_actions.append(download_button, send_button);
                table_row.append(name_element, last_name_element, row_actions);
                table.append(table_row);
            });
        });
    }
}
var image_element = document.getElementById("image_element");
let img;
let url;

function generatePDF_download(first_name, last_name) {
    if (test_fields()) {
        window.jsPDF = window.jspdf.jsPDF;
        const doc = new jsPDF({
            orientation: "landscape",
            unit: "mm",
            format: "a4",
        });
        const reader = new FileReader();
        reader.addEventListener(
            "load",
            () => {
                img = reader.result;
                doc.addImage(img, "JPEG", 0, 0, 297, 210);
                doc.addImage(fablab_url, "JPEG", 128.5, 3, 40, 40)
                // type d'attestation
                type_text = type_input.value.toUpperCase()
                doc.setFontSize(30)
                var textWidth = doc.getStringUnitWidth(type_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                var textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 60, type_text);
                // line one
                doc.setLineWidth(.6)
                doc.line(40, 70, 257, 70)
                // CELA CERTIFIE QUE
                certifie_text = "CELA CERTIFIE QUE"
                doc.setFontSize(20)
                textWidth = doc.getStringUnitWidth(certifie_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 85, certifie_text);
                // name_text
                name_text = first_name + " " + last_name
                doc.setFontSize(60)
                textWidth = doc.getStringUnitWidth(name_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 120, name_text);
                // description
                description = desc_input.value
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(description) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 140, description);
                
                // image two
                doc.addImage(greenlab_url, "JPEG", 128.5, 145, 40, 40)
                // line two
                doc.setLineWidth(.6)
                doc.line(120, 175, 177, 175)
                
                // manager
                manager_name = "AFASS AYOUB"
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(manager_name) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 185, manager_name);
                // manager text
                manager_text = "Manager de GreenLab"
                doc.setFontSize(13)
                textWidth = doc.getStringUnitWidth(manager_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 192, manager_text);
                doc.save(first_name + "_" + last_name + ".pdf");
            },
            false
        );
        if (bg_input.files[0]) {
            reader.readAsDataURL(bg_input.files[0]);
        }
    }
}

function send_solo_email(first_name, last_name, email) {
    if (test_fields()) {
        // generating database64 pdf
        window.jsPDF = window.jspdf.jsPDF;
        const doc = new jsPDF({
            orientation: "landscape",
            unit: "mm",
            format: "a4",
        });
        const reader = new FileReader();
        reader.addEventListener(
            "load",
            () => {
                img = reader.result;
                doc.addImage(img, "JPEG", 0, 0, 297, 210);
                doc.addImage(fablab_url, "JPEG", 128.5, 3, 40, 40)
                // type d'attestation
                type_text = type_input.value.toUpperCase()
                doc.setFontSize(30)
                var textWidth = doc.getStringUnitWidth(type_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                var textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 60, type_text);
                // line one
                doc.setLineWidth(.6)
                doc.line(40, 70, 257, 70)
                // CELA CERTIFIE QUE
                certifie_text = "CELA CERTIFIE QUE"
                doc.setFontSize(20)
                textWidth = doc.getStringUnitWidth(certifie_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 85, certifie_text);
                // name_text
                name_text = first_name + " " + last_name
                doc.setFontSize(60)
                textWidth = doc.getStringUnitWidth(name_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 120, name_text);
                // description
                description = desc_input.value
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(description) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 140, description);
                
                // image two
                doc.addImage(greenlab_url, "JPEG", 128.5, 145, 40, 40)
                // line two
                doc.setLineWidth(.6)
                doc.line(120, 175, 177, 175)
                
                // manager
                manager_name = "AFASS AYOUB"
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(manager_name) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 185, manager_name);
                // manager text
                manager_text = "Manager de GreenLab"
                doc.setFontSize(13)
                textWidth = doc.getStringUnitWidth(manager_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 192, manager_text);
                url = "data:application/pdf;base64," + btoa(doc.output());
                // then send email
                Email.send({
                    Host: "smtp.gmail.com",
                    Username: "certif.competition@gmail.com",
                    From: "certif.competition@gmail.com",
                    Password :'utnaxbmkfiizzqdl',
                    To: email,
                    Subject: "Certificat",
                    Body: document.getElementById("email-content").innerHTML,
                    Attachments: [
                        {
                            name: first_name + "_" + last_name + ".pdf",
                            data: url,
                        },
                    ],
                }).then((message) => {
                    console.log(message);
                    if (message == "OK")
                        show_message("Email sent successfully", "alert-success");
                    else show_message("Email not sent", "alert-danger");
                });
            },
            false
        );
        if (bg_input.files[0]) {
            reader.readAsDataURL(bg_input.files[0]);
        }
    }
}

function preview() {
    if (test_fields()) {
        if (iframe.hidden){
            preview_button.textContent = 'hide preview'
            iframe.hidden = false
        }else{
            preview_button.textContent = 'see preview'
            iframe.hidden = true

        }
        window.jsPDF = window.jspdf.jsPDF;
        const doc = new jsPDF({
            orientation: "landscape",
            unit: "mm",
            format: "a4",
        });
        const reader = new FileReader();
        reader.addEventListener(
            "load",
            () => {
                img = reader.result;
                doc.addImage(img, "JPEG", 0, 0, 297, 210);
                doc.addImage(fablab_url, "JPEG", 128.5, 3, 40, 40)
                // type d'attestation
                type_text = type_input.value.toUpperCase()
                doc.setFontSize(30)
                var textWidth = doc.getStringUnitWidth(type_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                var textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 60, type_text);
                // line one
                doc.setLineWidth(.6)
                doc.line(40, 70, 257, 70)
                // CELA CERTIFIE QUE
                certifie_text = "CELA CERTIFIE QUE"
                doc.setFontSize(20)
                textWidth = doc.getStringUnitWidth(certifie_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 85, certifie_text);
                // name_text
                name_text = "NOM PRNOM"
                doc.setFontSize(60)
                textWidth = doc.getStringUnitWidth(name_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 120, name_text);
                // description
                description = desc_input.value
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(description) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 140, description);

                // date
                date = date_input.value
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(date) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 150, date);
                
                // image two
                doc.addImage(greenlab_url, "JPEG", 128.5, 145, 40, 40)
                // line two
                doc.setLineWidth(.6)
                doc.line(120, 175, 177, 175)
                
                // manager
                manager_name = "AFASS AYOUB"
                doc.setFontSize(16)
                textWidth = doc.getStringUnitWidth(manager_name) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("primary-color").value)
                doc.text(textOffset, 185, manager_name);
                // manager text
                manager_text = "Manager de GreenLab"
                doc.setFontSize(13)
                textWidth = doc.getStringUnitWidth(manager_text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                textOffset = (doc.internal.pageSize.width - textWidth) / 2;
                doc.setTextColor(document.getElementById("secondary-color").value)
                doc.text(textOffset, 192, manager_text);

                url = "data:application/pdf;base64," + btoa(doc.output());
                var iframe = document.getElementById("iframe");
                iframe.src = url;
            },
            false
        );
        if (bg_input.files[0]) {
            reader.readAsDataURL(bg_input.files[0]);
        }
    }
}
