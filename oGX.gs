var today = new Date(); // Get the current date
var today = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");
var startDate = "2024/01/01" //Utilities.formatDate(new Date(today), "GMT+7", "dd/MM/yyyy"); //YYYY/MM/DD
var endDate = today
var sheetname = 'DATA STORAGE'
    var sheet_headers = [
    "ID",
    "EP Name",
    "Phone Number",
    "EP ID",
    "OP ID",
    "Home LC",
    "Home MC",
    "LC Alignment", /// difference from iCX
    "Title",
    "Host MC",
    "Host LC",
    "Product",
    "Status",
    "Signed Up At",
    "Applied At",
    "Accepted At",
    "Approved At",
    "Break Approve At",
    "Realized At",
    "Finished At",
    "Completed At",
    "Rejected Reason",
    "Email", 
    "CV Links",
    "Slot Start Date", 
    "Slot End Date",
    "Duration", 
    "Work Field",
    "NPS",
    "SDG",
    "SDG target",
    "SDG Index"
  ];

var accessToken = 'YOUR_ACCESS_TOKEN'; // Your access token
function getTotalPages() {
  // Define your GraphQL API endpoint and query for getting total pages
  var graphql_endpoint = 'https://gis-api.aiesec.org/graphql';
  var perPage = 1;  // Set per_page to 1 as we only need to get the total count
  var totalPagesQuery = `query GetTotalPages {
    allOpportunityApplication(
      per_page: ${perPage}
      filters: {
        person_home_mc: 504
        created_at:{
            from: "${startDate}" 
            to:"${endDate}"
            }
      }
    ) {
      paging {
        total_pages
      }
    }
  }`;

  try {
    var response = UrlFetchApp.fetch(graphql_endpoint, {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': accessToken,  // token eka
      },
      'payload': JSON.stringify({ 'query': totalPagesQuery }),
      'muteHttpExceptions': true
    });

    if (response.getResponseCode() == 200) {
      var data = JSON.parse(response.getContentText());
      Logger.log(data)

      if (data['data'] && data['data']['allOpportunityApplication']) {
        var total_pages = data['data']['allOpportunityApplication']['paging']['total_pages'];
        Logger.log("Total Pages: " + total_pages);
      } else {
        Logger.log("No 'allOpportunityApplication' property in the response for total pages.");
      }
    } else {
      Logger.log("GraphQL request failed with status code: " + response.getResponseCode() + " for total pages.");
    }
  } catch (error) {
    Logger.log("An error occurred: " + error + " for total pages.");
  }
}

// Call the function to log the total pages to the console
getTotalPages();


function convertToShortDate2(dateTimeStr) {
  // Parse the date-time string
  var dateTime = new Date(dateTimeStr);

  // Get the date components
  var year = dateTime.getFullYear();
  var month = ('0' + (dateTime.getMonth() + 1)).slice(-2); // Months are zero-indexed
  var day = ('0' + dateTime.getDate()).slice(-2);

  // Construct the short date string (YYYY-MM-DD)
  var shortDate = day + '-' + month + '-' + year;
  return shortDate;
}


function convertDatesToShortFormat2(nameSheet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  var range = sheet.getRange("N2:U" + sheet.getLastRow()); // Adjust range as per your column indices
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] !== "") {
        values[i][j] = convertToShortDate2(values[i][j]);
      }
    }
  }

  range.setValues(values);
}


function iCXExtract(fr, t, nameSheet) {
  // Define your GraphQL API endpoint and query
  var graphql_endpoint = 'https://gis-api.aiesec.org/graphql';
  var perPage = 1000;  // Number of items per page
  var page = 1;
  var allApplications = [];

  // Defining Google Sheets spreadsheet and the sheet where you want to write the data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(nameSheet); // sheet eke nama

  // Define headers for the Google Sheets

  // Send the GraphQL request with headers in a loop until all pages are retrieved
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': accessToken,  // Your token
  };
  while (true) {
    var graphql_query = `query{
	allOpportunityApplication(
		page: ${page}
		per_page: ${perPage}
		filters:{
			person_home_mc: 504
			created_at:{from:"${fr}",to:"${t}"}
		}
	)
	{
		data {
			id
      person{
				email
        signed_up_at: created_at
        id
				cv_url
				full_name
				home_lc{
					name
				}
        home_mc{
          name
        }
				lc_alignment {
					id
					label
          keywords
				}
        contact_detail{
					email
					facebook
					instagram
					linkedin
					twitter
					phone
				}
				updated_at
			}
			current_status
			cv{
				id
				url
			}
			created_at
			date_matched
			date_approved
			date_approval_broken
			date_realized
      experience_end_date
      nps_response_completed_at
			nps_grade
      rejection_reason {
				meta
				name
			}
      home_mc {
          name
        }
        host_lc {
          name
        }
      host_lc_name
			status
      opportunity{
				id
        organisation{
					name
				}
        title
        opportunity_duration_type {
					duration_type
				}
        sub_product {
					name
				}
				sdg_info {
					id
          sdg_target {
						id
						goal_index
						target
						target_id
						target_index
					}
				}
        programme{
          id
          short_name_display
        }
			}
      slot{
        start_date
        end_date
      }
		}
    paging {
          current_page
          total_items
          total_pages
    }
	}
}`;

    try {
      var response = UrlFetchApp.fetch(graphql_endpoint, {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify({ 'query': graphql_query }),
        'muteHttpExceptions': true,
      });

      //Logger.log(response.getContentText())

      if (response.getResponseCode() == 200) {
        var data = JSON.parse(response.getContentText());

        allApplications = allApplications.filter(application => {
          return new Date(application.created_at) >= new Date(startDate);
        });

        if (data['data'] && data['data']['allOpportunityApplication']) {
          var applications = data['data']['allOpportunityApplication']['data'];

          allApplications = allApplications.concat(applications);

          // Check if there are more pages to retrieve
          var currentPage = data['data']['allOpportunityApplication']['paging']['current_page'];
          var totalPages = data['data']['allOpportunityApplication']['paging']['total_pages'];

          if (currentPage < totalPages) {
            page++;  // Move to the next page
          } else {
            break;  // Exit the loop if all pages are retrieved
          }
        } else {
          Logger.log("No 'allOpportunityApplication' property in the response.");
          //Logger.log("Full response: " + JSON.stringify(data));  // Log the full response for debugging
          break;
        }
      } else {
        Logger.log("GraphQL request failed with status code: " + response.getResponseCode());
        Logger.log("Response content: " + response.getContentText());  // Log the response content for debugging
        break;  // Exit the loop in case of an error
      }
    } catch (error) {
      Logger.log("An error occurred: " + error);
      break;  // Exit the loop in case of an error
    }
  }

  // Clear existing data in the sheet
  //sheet.clearContents();

  // Write the headers to the sheet
  //sheet.getRange(1, 1, 1, sheet_headers.length).setValues([sheet_headers]);

  // Write the data to the sheet if there is any
  if (allApplications.length > 0) {

    var rows = [];
    for (var i = 0; i<allApplications.length; i++) {
      var application = allApplications[i];

      // Handle potential null values with null checks
      try {
        var row = [
          application['id'] ? application['id'] : '',
          application['person'] ? application['person']['full_name'] : '',
          application['person'] && application['person']['contact_detail'] ? application['person']['contact_detail']['phone'] : '',
          application['person'] ? application['person']['id'] : '',
          application['opportunity'] ? application['opportunity']['id'] : '',
          application['person'] ? application['person']['home_lc'] ? application['person']['home_lc']['name'] : '' : '',
          application['person'] ? application['person']['home_mc'] ? application['person']['home_mc']['name'] : '' : '',
          application['person'] ? application['person']['lc_alignment'] ? application['person']['lc_alignment']['keywords'] : '' : '',
          application['opportunity'] ? application['opportunity']['title'] : '',
          application['home_mc'] ? application['home_mc']['name'] : '',
          application['host_lc'] ? application['host_lc']['name'] : '',
          application['opportunity'] ? application['opportunity']['programme'] ? application['opportunity']['programme']['short_name_display'] : '' : '',
          application['status'] ? application['status'] : '',
          application['person'] ? application['person']['signed_up_at'] : '',
          application['created_at'] ? application['created_at'] : '',
          application['date_matched'] ? application['date_matched'] : '',
          application['date_approved'] ? application['date_approved'] : '',
          application['date_approval_broken'] ? application['date_approval_broken'] : '',
          application['date_realized'] ? application['date_realized'] : '',
          application['experience_end_date'] ? application['experience_end_date'] : '',
          application['nps_response_completed_at'] ? application['nps_response_completed_at'] : '',
          application['rejection_reason'] ? application['rejection_reason']['name'] : '', //không rút gọn hàng này
          application['person'] && application['person']['contact_detail'] ? application['person']['contact_detail']['email'] : '',
          application['person'] ? application['person']['cv_url'] : '',
          application['slot'] ? application['slot']['start_date'] : '',
          application['slot'] ? application['slot']['end_date'] : '',
          application['opportunity'] ? application['opportunity']['opportunity_duration_type'] ? application['opportunity']['opportunity_duration_type']['duration_type'] : '' : '', //không rút gọn hàng này
          application['opportunity'] ? application['opportunity']['sub_product'] ? application['opportunity']['sub_product']['name'] : '' : '', //không rút gọn hàng này
          application['nps_grade'] ? application['nps_grade'] : '',
          application['opportunity'] ? application['opportunity']['sdg_info'] ? application['opportunity']['sdg_info']['sdg_target'] ? application['opportunity']['sdg_info']['sdg_target']['goal_index'] : '' : '' : '',
          application['opportunity'] ? application['opportunity']['sdg_info'] ? application['opportunity']['sdg_info']['sdg_target'] ? application['opportunity']['sdg_info']['sdg_target']['target'] : '' : '' : '',
          application['opportunity'] ? application['opportunity']['sdg_info'] ? application['opportunity']['sdg_info']['sdg_target'] ? application['opportunity']['sdg_info']['sdg_target']['target_index'] : '' : '' : ''
        ];
        rows.push(row);
      }
      catch (error) {
        console.log(error.message, error.stack);
        continue;
      }
    }

    // Write all the rows to the sheet
    sheet.insertRowsBefore(2, rows.length)
    sheet.getRange(2, 1, rows.length, sheet_headers.length).setValues(rows);
  } else {
    Logger.log("No data to write to the sheet.");
  }
}

function exactYear() {


  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname)
  sheet.clearContents()
  sheet.deleteRows(3, sheet.getMaxRows()-3+1)
  sheet.insertRowsAfter(2, 10) /// thêm dòng để khỏi lỗi
  sheet.getRange(1, 1, 1, sheet_headers.length).setValues([sheet_headers]);
  var start = new Date(startDate) //YYYY-MM-DD
  var today = new Date()
  var end = today
  for (var year = start.getFullYear(); year<=end .getFullYear(); year++) {
    var tmp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("[FOR DUPLICATE]")
    var startYearDate
    var endYearDate
    if (start.getFullYear() == year) startYearDate = start
    else {
      startYearDate = new Date()
      startYearDate.setYear(year)
      startYearDate.setMonth(0)
      startYearDate.setDate(1)
    }
    if (end.getFullYear() == year) endYearDate = end
    else {
      endYearDate = new Date()
      endYearDate.setYear(year)
      endYearDate.setMonth(11)
      endYearDate.setDate(31)
    }
    var startS = Utilities.formatDate(startYearDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'00:00:00");
    var endS = Utilities.formatDate(endYearDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'23:59:59");
    Logger.log(startS)
    Logger.log(endS)
    //iCXExtract(startS, endS, String(year))
    iCXExtract(startS, endS, sheetname)
  }
  convertDatesToShortFormat2(sheetname);
  //sheet.getDataRange().getValues();
}

