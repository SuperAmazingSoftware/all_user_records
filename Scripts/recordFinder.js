// JS to handle all Record Finder Page Events
var jq = jQuery.noConflict();
var entityList = [];
var userList = [];
var attributeList = [];

var globalContextObj = {};
window.orgInfo = {
  serviceRootUrl: "",
  _getServiceRootUrl: function () {
    return window.orgInfo.serviceRootUrl;
  },
  _initializeServiceRootUrl: function (globalContext) {
    if (globalContext !== null) {
      let clientUrl = globalContext.getClientUrl();
      let fullVer = globalContext.getVersion();
      let version = fullVer.substring(3, fullVer.indexOf(".") - 1);

      window.orgInfo.serviceRootUrl = clientUrl.concat("/api/data/v", version);
    }
  },
};

$(document).ready(function () {
  globalContextObj = getGlobalContext();
  window.orgInfo._initializeServiceRootUrl(globalContextObj);

  if (globalContextObj !== null)
    $("#submit").click(() =>
      fetchRecordsByUser(
        getSelectedUser(),
        getSetlectedEntity(),
        generateResultTable
      )
    );

    $("#select-entity").change((event) => {
      console.log($(event.target).children("option:selected").val());
      fetchAttributeList(entityList[0].MetadataId, generateAttributeList);
    })

  $("#loadOrgData").click(function () {
    $("#select-entity").empty();

    let entityData;
    let userData;

    let xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        entityData = JSON.parse(this.responseText);
        entityData = entityData.value;

        $(entityData).each(function (index, element) {
          entityList.push(element);
        });

        for (var i = 0; i <= entityList.length; i++) {
          if (typeof entityList[i] !== "undefined" && typeof entityList[i].EntitySetName !== "undefined") {
            $("#select-entity").append(new Option(entityList[i].DisplayName.LocalizedLabels[0].Label,entityList[i].MetadataId));
          }
          // make sure to create the JSON object, push it to the entityName Array, iterate the array, extract the values, and pass them to each new Option object)
          // $("#select-entity").append(new Option(entityList[i].DisplayName.LocalizedLabels[0].Label, entityList[i].EntitySetName));
        }        
        entityList.sort();

      }
    };
    xhr.open(
      "GET",
      window.orgInfo._getServiceRootUrl() +
        "/EntityDefinitions?$select=MetadataId,Attributes,DisplayName,EntitySetName&$filter=OwnershipType eq Microsoft.Dynamics.CRM.OwnershipTypes'UserOwned'"
    );
    xhr.send();

    getSystemUsers(populateUserNames);
  });
});

function getSelectedAttributes(){
  return ["_createdby_value",
  "createdon",
  "_modifiedby_value",
  "modifiedon",
  "_owninguser_value"]
}

function getSelectedUser() {
  return { systemUserId: $("#select-user").val() };
}

function getSetlectedEntity() {
  let entity = entityList.find(
    (entity) => entity.MetadataId == $("#select-entity").val()
  );
  return entity.EntitySetName;
}

function getGlobalContext() {
  if (parent.Xrm.Utility.getGlobalContext() !== null)
    return parent.Xrm.Utility.getGlobalContext();
}

function getUserOwnedEntities() {
  let entityData;
  let xhr = new XMLHttpRequest();
  xhr.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      entityData = JSON.parse(this.responseText);
      getEntityNames(entityData);
    }
  };
  xhr.open(
    "GET",
    window.orgInfo._getServiceRootUrl() +
      "/EntityDefinitions?$select=DisplayName,EntitySetName&$filter=OwnershipType eq Microsoft.Dynamics.CRM.OwnershipTypes'UserOwned'"
  );
  xhr.send();
  // Next get all the entity names from the json array
}

function getSystemUsers(callback) {
  let systemUsers;
  let xhr = new XMLHttpRequest();
  xhr.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      systemUsers = JSON.parse(this.responseText);
      ingestUsers(systemUsers);
      callback(systemUsers);
    }
  };
  xhr.open("GET", window.orgInfo._getServiceRootUrl() + "/systemusers");
  xhr.send();
}
// collects
function getEntityNames(entityData) {
  let value = entityData["value"];
  for (keys in value) {
    let logicalName = value[keys]["LogicalName"];
    entityList.push(logicalName);
  }
}

function ingestUsers(systemUsers) {
  let value = systemUsers["value"];
  for (keys in value) {
    let userID = value[keys]["systemuserid"];
    let userName = value[keys]["fullname"];
    let userObject = { systemUserId: `${userID}`, userName: `${userName}` };
    userList.push(userObject);
  }
  userList.sort();
}

function populateUserNames(userData) {
  for (var o = 0; o < userData.value.length; o++) {
    $("#select-user").append(
      new Option(userList[o].userName, userList[o].systemUserId)
    );
  }
}

function fetchRecordsByUser(systemUser, entitySetName, callback) {
  // Fetch XML query
  // Create request
  var userId = systemUser.systemUserId;
  var columnSetString = "";
  getSelectedAttributes().forEach((column) => (columnSetString += column + ","));
  columnSetString = columnSetString.substr(0, columnSetString.length - 1);

  var req = new XMLHttpRequest();
  req.open(
    "GET",
    window.orgInfo._getServiceRootUrl() +
      "/" +
      entitySetName +
      "?$select=" +
      columnSetString +
      "&$filter=_owninguser_value eq " +
      userId,
    true
  );
  req.setRequestHeader("OData-MaxVersion", "4.0");
  req.setRequestHeader("OData-Version", "4.0");
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
  req.onreadystatechange = function () {
    if (this.readyState === 4) {
      req.onreadystatechange = null;
      if (this.status === 200) {
        var results = JSON.parse(this.response);
        for (var i = 0; i < results.value.length; i++) {
          var _createdby_value = results.value[i]["_createdby_value"];
          var _createdby_value_formatted =
            results.value[i][
              "_createdby_value@OData.Community.Display.V1.FormattedValue"
            ];
          var _createdby_value_lookuplogicalname =
            results.value[i][
              "_createdby_value@Microsoft.Dynamics.CRM.lookuplogicalname"
            ];
          var createdon = results.value[i]["createdon"];
          var _modifiedby_value = results.value[i]["_modifiedby_value"];
          var _modifiedby_value_formatted =
            results.value[i][
              "_modifiedby_value@OData.Community.Display.V1.FormattedValue"
            ];
          var _modifiedby_value_lookuplogicalname =
            results.value[i][
              "_modifiedby_value@Microsoft.Dynamics.CRM.lookuplogicalname"
            ];
          var modifiedon = results.value[i]["modifiedon"];
          var _owninguser_value = results.value[i]["_owninguser_value"];
          var _owninguser_value_formatted =
            results.value[i][
              "_owninguser_value@OData.Community.Display.V1.FormattedValue"
            ];
          var _owninguser_value_lookuplogicalname =
            results.value[i][
              "_owninguser_value@Microsoft.Dynamics.CRM.lookuplogicalname"
            ];
        }
        callback(results.value);
      } else {
        console.log(this.statusText);
      }
    }
  };
  req.send();
}

function generateResultTable(data) {
  var target = "#results-view";
  var innerHtml = "";
  innerHtml += '<table class="table table-bordered">';
  innerHtml += "<thead>";
  innerHtml += '<tr scope="row">';

  $(getSelectedAttributes()).each(function (index, element) {
    innerHtml += '<th scope="col">' + element + "</th>";
  });

  

  innerHtml += "/<tr>";
  innerHtml += "</thead>";

  innerHtml += '<tbody class="table-striped table-hover">';

  for (i = 0; i < data.length; i++) {
    innerHtml += '<tr scope="row">';
    for (o = 0; o < columnSetArray.length; o++) {
      value = data[i][columnSetArray[o]];
      innerHtml += "<td>" + value + "</td>";
    }
    innerHtml += "</tr>";
  }
  innerHtml += "</tbody>";
  innerHtml += "</table>";

  $(target).html(innerHtml);
}

function fetchAttributeList(metadataId,callback){
  let xhr = new XMLHttpRequest();
  xhr.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        attributeList = JSON.parse(this.responseText);
        attributeList = attributeList.Attributes;
        callback(attributeList);
      }
  } 
  xhr.open("GET",window.orgInfo._getServiceRootUrl() + `/EntityDefinitions(${metadataId})?$select=MetadataId&$expand=Attributes($select=LogicalName,DisplayName)` ); 
  xhr.send();
}

function generateAttributeList(attributes){
  $("#select-attributes").empty();
  $(attributes).each(function (index, attribute) {
    if(attribute.DisplayName.LocalizedLabels && attribute.DisplayName.LocalizedLabels.length > 0 ){
      $("#select-attributes").append(new Option(attribute.DisplayName.LocalizedLabels[0].Label,attribute.LogicalName));
    }
  });
}


