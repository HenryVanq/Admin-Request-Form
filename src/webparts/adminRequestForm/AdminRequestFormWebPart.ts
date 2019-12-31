import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
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

var html = ''

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
        <input type="text" minlength="10" maxlength="10" class="form-control" id="date" name="txtDate" placeholder="${strings.DescriptionFieldLabelDate}" required="true">
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
        <input minlength="10" maxlength="10" class="form-control" id="phoneNumber" placeholder="${strings.DescriptionFieldLabelPhoneNumber}" required="true">
        <small id="phoneHelpBlock" class="form-text text-muted"></small>
        </div>
      <div class="form-group col-md-6">
        <label for="inputCity"> <h6> ${strings.DescriptionFieldLabelEmail} *  </h6> </label>
        <input type="email" class="form-control" id="email" placeholder="${strings.DescriptionFieldLabelEmail}" required="true">
      </div>
    </div>

    <div class="form-group">
    <label for="exampleFormControlTextarea1"><h6> ${strings.DescriptionFieldLabelReason} * </h6> </label>
    <textarea maxlength="2000"class="form-control" id="reason" rows="2" placeholder="${strings.DescriptionFieldLabelReason}" required="true"></textarea>
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

</form>`;

    (<any>$("#date")).datepicker(
      {
        changeMonth: true,
        changeYear: true,
        dateFormat: "dd-mm-yy"
      }
    );

    $('#submit').on("click", () => {
      this.postingData();
    })

    $('#cancel').on('click', () => {
      if (confirm("Ακύρωσης αιτήματος?")) {
        window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing');
      }
    })

    sp.web.lists.getByTitle("Department").items.get().then((data) => {
      $("#selectteam").empty();
      data.map((item) => {
        $('#selectteam').append('<option>' + item.NameDepartment + '</option>')
      })
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
    });
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    return super.onInit()
  }

  protected async postingData() {
    const inputRequest = $('#request').val();
    const inputSelectDepartment = $('#selectteam').val();
    const inputRefNumberIn = $('#refNumberIn').val();
    const date = $("#date").val()
    const inputFullName = $('#fullname').val();
    const inputOrganization = $('#organization').val();
    const inputEmail = $('#email').val();
    const inputPhoneNumber = $('#phoneNumber').val();
    const inputReason = $('#reason').val();
    const attachedFile = $('#fileUploadInput').val();

    if (inputRequest === ""
      || inputSelectDepartment === ""
      || inputRefNumberIn === ""
      || (date === "" || date.toString().length != 10)
      || inputFullName === ""
      || inputOrganization === ""
      || inputEmail === ""
      || inputPhoneNumber === ""
      || inputReason === "") {

      return false
    }


    if (!this.validateEmail(inputEmail)) {
      return false
    }

    if (this.checkingLength() === false) {
      $('#phoneHelpBlock').html('<p>  Το τηλέφωνο επικοινωνίας πρέπει να περιλαμβάνει μόνο αριθμούς</p>')
    }

    var num = new Number(inputPhoneNumber);
    var strPhoneNumber = num.toString();

    if (strPhoneNumber.length != 10) {
      return false
    }


    await sp.web.lists.getByTitle("Department").items.get().then((item: any) => {
      item.map((data) => {
        if (inputSelectDepartment != data.NameDepartment) {
          return false
        }

        sp.web.lists.getByTitle('ExternalRequest').items.add({
          Request: inputRequest,
          Date: date.toString(),
          Fullname: inputFullName,
          Organization: inputOrganization,
          Email: inputEmail,
          PhoneNumber: inputPhoneNumber,
          Reason: inputReason,
          ReferenceNumberIn: inputRefNumberIn,
          Department: inputSelectDepartment,
          Available: "yes",
          DepartmenPhone: data.PhoneDepartment,
          emailDepartment: data.email

        }).then((result) => {
          for (let i in result) {
            if (attachedFile === "") {
              return false
            }
            return this.uploadingFileEventHandlers('ExternalRequest', result[i].ID);
          }

        }).catch(e => console.log(e))
      })
    })

    await this.hidingForm(inputRefNumberIn)

  }


  protected hidingForm(inputRefNumberIn) {
    $("#form").hide();
    $('#submit').prop("disabled", true);
    $('#divShow').append(`
          <div class="card"> 
            <div class="card-body">
              <h6 class="card-title">Η υποβολή του αιτήματός σας με αναγνωριστικό «${inputRefNumberIn}» ολοκληρώθηκε επιτυχώς. Θα ειδοποιηθείτε στη ηλεκτρονική διεύθυνση αλληλογραφίας που δηλώσατε για τη λήψη των στοιχείων μετά από την επεξεργασία του αιτήματός σας από το αρμόδιο τμήμα </h6>
            </div>
          </div> 
              <br>
          <button id="ok" class="btn btn-light btn-block border"> <h5> Μετάβαση στη σελίδα διαχείρισης <h5> </button> <br>`);

    $('#ok').on('click', () => {
      window.location.replace('https://idikagr.sharepoint.com/sites/ExternalSharing');
    })
  }

  protected checkingLength() {
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
