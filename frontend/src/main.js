import './style.css';
import './app.css';

import logo from './assets/images/logo-universal.png';
import {SelectFiles, ProcessFiles} from '../wailsjs/go/main/App';

document.querySelector('#app').innerHTML = `
    <img id="logo" class="logo">
        <h1>Convert EXCEL to HTML</h1>
        <p>Support for generating HTML for BPS static tables</p>
        <div class="input-box" id="input">
            <button class="input" id="selectFileButton"onclick="selectFiles()">Add Excel Files</button>
            <div class="select-files"id="fileList"></div>
            <button class="btn" id="submitFilesButton"onclick="processFiles()">Convert</button>
        </div>
        <div class="log-result"id="log"></div>
    </div>
`;
document.getElementById('logo').src = logo;
var fileList = document.getElementById('fileList');
var log = document.getElementById('log');

const selectedFiles = [];

window.selectFiles = function () {
    // Call App.Select(name)
    try {
        SelectFiles()
        .then((result) => {
            //selectedFiles.push(result);
            result.forEach(function(item, index) {
                var index = selectedFiles.indexOf(item);
                if (index > -1) { // only when item is found
                    return;
                }
                selectedFiles.push(item);
                const listItem = document.createElement('div');
                listItem.setAttribute('title', item);
                var filename = item.replace(/^.*[\\/]/, '');
                listItem.innerHTML = '<span>'+filename+'</span>' + '<button class="x"data-path=\"'+item+'\"onclick="removeFile(this)">x</button>';
                fileList.appendChild(listItem);
            });              
            
        })
        .catch((err) => {
            console.error(err);
        });
    } catch (err) {
        console.error(err);
    }
};

window.processFiles = function () {
    if (selectedFiles.length === 0) {
        alert('No files selected');
        return;
    }
    // Call App.Select(name)
    try {
        ProcessFiles(selectedFiles)
        .then((result) => {
            resetForm();
            log.innerHTML ="";
            result.forEach(function(item, index) {
                const listItem = document.createElement('p');
                if (item.indexOf('success :') != 0) {
                    listItem.setAttribute('class','red');
                }
                listItem.textContent = item;
                log.appendChild(listItem);
            }); 
        })
        .catch((err) => {
            console.error(err);
        });
    } catch (err) {
        console.error(err);
        alert(error);
    }
};
window.removeFile = function (e) {
    var fpath = e.getAttribute("data-path");
    e.parentNode.parentNode.removeChild(e.parentNode);
    var index = selectedFiles.indexOf(fpath);
    if (index > -1) { // only splice array when item is found
        selectedFiles.splice(index, 1); // 2nd parameter means remove one item only
    }
};

function resetForm(){
    selectedFiles.length = 0;
    fileList.innerHTML = ""
}




