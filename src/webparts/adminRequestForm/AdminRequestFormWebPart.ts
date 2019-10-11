import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'AdminRequestFormWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';

import { sp, ItemAddResult } from "@pnp/sp";

import * as $ from 'jquery';

require('./css/jquery-ui.css');
require('jqueryui');

let cssUrl = 'https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css';
SPComponentLoader.loadCss(cssUrl)

export interface IAdminRequestFormWebPartProps {
  description: string;
  fileName: string;
  fullName: string;
  organization: string;
  phoneNumber: string;
  email: string;
  reason: string;
  status: string;
  downloaded: string;
  statusPending: string;
  statusApproved: string;
  statusRejected: string;
}

export default class AdminRequestFormWebPart extends BaseClientSideWebPart<IAdminRequestFormWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
     
      
    <div class="card text-center bg-info mb-3">
      <div class="card-header"> <h3 id="title" class="text-white">Αίτημα Παροχής Στοιχείων </h3> </div>
    </div>

    <div class="ol text-center" id="divShow" class="form-group"> </div>

    <form id="form" onsubmit="return false" autocomplete="off">

    <div class="form-row">
      <div class="form-group col-md-6">
        <label for="inputEmail4"> <h6> ${strings.DescriptionFieldLabelRequest} * </h6> </label>
        <input maxlength="255" type="text" class="form-control" id="request" placeholder="${strings.DescriptionFieldLabelRequest}" required="true" autocomplete="off">
      </div>
      
      <div class="form-group col-md-6">
      <label for="inputEmail4"> <h6>  Επιλογή Τμήματος * </h6> </label>
        <select id="selectteam" class="form-control" placeholder="test">
          <option selected>  </option>
          <option> Γενικό </option>
          <option> Τμήμα Α </option>
          <option> Τμήμα Β </option>
          <option> Τμήμα Γ </option>
          <option> Τμήμα Δ </option>
        </select>
      </div>
    </div>

    <div class="form-row" >
      <div class="form-group col-md-6">
        <label for="inputEmail4"> <h6> ${strings.DescriptionFieldLabelRefNumberIn} * </h6> </label>
        <input maxlength="255" type="text" class="form-control" id="refNumberIn" placeholder="${strings.DescriptionFieldLabelRefNumberIn}" required="true" autocomplete="off">
      </div>
      <div class="form-group col-md-6">
        <label for="inputEmail4"> <h6> ${strings.DescriptionFieldLabelDate} * </h6> </label>
        <input type="text" class="form-control" id="date" name="txtDate" placeholder="${strings.DescriptionFieldLabelDate}" required="true">
      </div>
    </div>

    <div class="form-row" >
      <div class="form-group col-md-6"">
        <label for="inputPassword4"> <h6> ${strings.DescriptionFieldLabelFullName} *  </h6> </label>
        <input maxlength="255"type="text" class="form-control" id="fullname" placeholder="${strings.DescriptionFieldLabelFullName}" required="true" >
      </div>
      <div class="form-group col-md-6"">
        <label for="inputAddress"> <h6> ${strings.DescriptionFieldLabelOrganization}  *  </h6> </label>
        <input maxlength="255" type="text" class="form-control" id="organization" placeholder="${strings.DescriptionFieldLabelOrganization}" required="true">
      </div>
    </div>

    <div class="form-row">
      <div class="form-group col-md-6">
        <label for="inputState"> <h6> ${strings.DescriptionFieldLabelPhoneNumber} *  </h6> </label>
        <input minlength="10" maxlength="10" type="tel" class="form-control" id="phoneNumber" placeholder="${strings.DescriptionFieldLabelPhoneNumber}" required="true">
      </div>
      <div class="form-group col-md-6">
        <label for="inputCity"> <h6> ${strings.DescriptionFieldLabelEmail} *  </h6> </label>
        <input type="email" class="form-control" id="email" placeholder="${strings.DescriptionFieldLabelEmail}" required="true">
      </div>
    </div>

    <div class="form-group">
    <label for="exampleFormControlTextarea1"><h6> ${strings.DescriptionFieldLabelReason} * </h6> </label>
    <textarea maxlength="2000"class="form-control" id="reason" rows="2" placeholder="${strings.DescriptionFieldLabelReason}"></textarea>
  </div>

  <div class="form-group">
    <label for="exampleFormControlTextarea1"><h6> Ανάρτηση Αρχείου </h6> </label>
    <input class="form-control"  type='file' id='fileUploadInput' name='myfile' />
    <button id="fileUpload" name="uFile" style='display:none'>upload</button>
  </div>

  <br>
  <div class="ol text-center">
    <button id="submit" type="submit" class="btn btn-dark btn-block"><h5> Αποστολή Αιτήματος </h5> </button>
    <button id="cancel" type="button" class="btn btn-light btn-block border"> <h5> Ακύρωση </h5> </button>
  </div>

</form>
        `;

    (<any>$("#date")).datepicker(
      {
        changeMonth: true,
        changeYear: true,
        dateFormat: "mm/dd/yy"
      }
    );

    $('#submit').on("click", () => {
      this.postingData('Requests');
    })

    $('#cancel').on('click', () => {
      if (confirm("Ακύρωσης αιτήματος?")) {

        window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing');

      }
    })
  }

  private uploadingFileEventHandlers(list, id): void {

    let fileUpload = document.getElementById("fileUpload")
    let test1 = document.getElementById("fileUploadInput")

    if (fileUpload) {
      this.uploadFiles(test1, list, id);
    }
  }

  private uploadFiles(fileUpload, list, id) {

    let file = fileUpload.files[0];
    let item = sp.web.lists.getByTitle(list).items.getById(id);

    item.attachmentFiles.add(file.name, file).then(v => {
      console.log(v)
    });
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    return super.onInit()
  }

  protected postingData(list: string) {

    const inputRequest = $('#request').val();
    const inputSelectDepartment = $('#selectteam').val();
    const inputRefNumberIn = $('#refNumberIn').val();
    const date = $("#date").val()
    const inputFullName = $('#fullname').val();
    const inputOrganization = $('#organization').val();
    const inputEmail = $('#email').val();
    const inputPhoneNumber = $('#phoneNumber').val();
    const inputReason = $('#reason').val();

    console.log("department: " + inputSelectDepartment)

    if (inputRequest != ""
      && inputSelectDepartment != ""
      && inputRefNumberIn != ""
      && date != ""
      && inputFullName != ""
      && inputOrganization != ""
      && inputEmail != ""
      && inputPhoneNumber != ""
      && inputReason != ""
      && this.validateEmail(inputEmail)
      && this.t() === true) {



      var num = new Number(inputPhoneNumber);
      var strPhoneNumber = num.toString();

      if (strPhoneNumber.length === 10) {
        try {
          // add an item to the list
          sp.web.lists.getByTitle(list).items.add({
            Request: inputRequest,
            RequestDate: date.toString(),
            Fullname: inputFullName,
            Organization: inputOrganization,
            Email: inputEmail,
            PhoneNumber: inputPhoneNumber,
            Reason: inputReason,
            ReferenceNumberIn: inputRefNumberIn,
            Department: inputSelectDepartment
          }).then((iar: ItemAddResult) => {
            console.log('iar:' + iar)
            for (let i in iar) {
              const input = $('#fileUploadInput').val();
              if (input === "") {
                return console.log('no file')
              }
              console.log('iar[i].ID: ' + iar[i].ID);
              console.log('list: ' + list)
              return this.uploadingFileEventHandlers(list, iar[i].ID);
            }
          }).then(() => {

            $("#form").hide();
            $('#submit').prop("disabled", true);

            $('#divShow').append(`
              <div class="card"> 
                <div class="card-body">
                  <h5 class="card-title">To αίτημα σας καταχωρήθηκε επιτυχώς </h5>
                </div>
              </div> 
              <br>
              <button id="ok" class="btn btn-light btn-block border"> <h5> Έξοδος <h5> </button> <br>
            `);

            $('#ok').on('click', () => {
              if (confirm("Έξοδος?")) {
                window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing');
              }
            })
          })
        } catch (err) {
          console.log(err)
        }
      } else {
        alert("To τηλέφωνο επικοινωνίας πρέπει να περιλαμβάνει 10 αριθμούς")
      }

    } else if (inputRequest != ""
      && inputSelectDepartment != ""
      && inputRefNumberIn != ""
      && date != ""
      && inputFullName != ""
      && inputOrganization != ""
      && inputEmail != ""
      && inputPhoneNumber != ""
      && inputReason != ""
      && this.validateEmail(inputEmail)
      && this.t() === false) {

      alert("To τηλέφωνο επικοινωνίας πρέπει να περιλαμβάνει μόνο αριθμούς")

    } else if (inputRequest != ""
      && inputSelectDepartment != ""
      && inputRefNumberIn != ""
      && date != ""
      && inputFullName != ""
      && inputOrganization != ""
      && inputEmail != ""
      && inputPhoneNumber != ""
      && inputReason != ""
      && !this.validateEmail(inputEmail)
      && this.t() === true) {
      alert("To Email που εισάγεται δεν είναι σωστό")
    } else if (inputRequest != ""
      && inputSelectDepartment != ""
      && inputRefNumberIn != ""
      && date != ""
      && inputFullName != ""
      && inputOrganization != ""
      && inputEmail != ""
      && inputPhoneNumber != ""
      && inputReason != ""
      && !this.validateEmail(inputEmail)
      && this.t() === false) {
      alert("To Email και τo τηλέφωνο επικοινωνίας δεν είναι σωστά")
    } else {

      alert("Πρέπει να συμπληρώστε όλα τα πεδία")
    }
  }

  protected t() {
    var x = document.forms["form"]["phoneNumber"].value;
    if (isNaN(x)) {
      return false;
    } else {
      return true;
    }
  }

  protected validateEmail(email) {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
