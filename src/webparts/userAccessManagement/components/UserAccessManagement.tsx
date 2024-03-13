import * as React from "react";
import {
  IUserAccessManagementProps,
  IUserAccessManagementState,
} from "./IUserAccessManagement";
import swal from "sweetalert2";
import "../assets/css/style.css";
import "bootstrap/dist/css/bootstrap.min.css";
import axios from "axios";
import "bootstrap-icons/font/bootstrap-icons.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "../assets/css/feather.css";
import * as $ from "jquery";
import * as moment from 'moment';
import { DatePicker } from 'antd';
import 'antd/dist/antd.css';
const { RangePicker } = DatePicker;


import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Select from "react-Select";
import "../assets/css/dataTables.bootstrap5.min.css";
import { Constants } from "./Constants/Constant";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as msal from "@azure/msal-browser";


// const msalConfig: msal.Configuration = {
//   auth: {
//     clientId: 'b9fd59c2-03b9-4540-9c4d-1a255a3889ea',
//     authority: 'https://login.microsoftonline.com/ff49c438-c469-4c10-96f6-61f54df41c9b'
//   },
//   cache: {
//     cacheLocation: "sessionStorage", // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
//     storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge

//   },
//   system: {
//     iframeHashTimeout: 10000,
//     loggerOptions: {
//       loggerCallback: (level, message, containsPii) => {
//         if (containsPii) {
//           return;
//         }
//         switch (level) {
//           case msal.LogLevel.Error:
//             console.error(message);
//             return;
//           case msal.LogLevel.Info:
//             console.info(message);
//             return;
//           case msal.LogLevel.Verbose:
//             console.debug(message);
//             return;
//           case msal.LogLevel.Warning:
//             console.warn(message);
//             return;
//         }
//       },
//     },
//   },
// };


// const msalInstance: msal.PublicClientApplication = new msal.PublicClientApplication(
//   msalConfig
// );
let currentAccount: msal.AccountInfo = null;
// const tokenrequest: any = {
//   scopes: ['api://b9fd59c2-03b9-4540-9c4d-1a255a3889ea/user_impersonation'],
//   //["Mail.Read"] ,//['api://c0f96cba-486b-4867-a2e2-e60ad996a3ed/hello'], ["Mail.Read"],
//   //scopes:["User.Read"],
//   account: currentAccount,
// };


export default class UserAccessManagement extends React.Component<
  IUserAccessManagementProps,
  IUserAccessManagementState
> {
  branch: any;
  constructor(props: any) {
    super(props);
    this.state = {
      employeeID: "",
      newOption: [],

      userName: "",
      branch: "",
      emailId: "",
      roles: [],
      datarecords: [],
      values: [],
      isActive: true,
      applicationSelectedValue: [],
      applicationSelectedRoles: [],
      Userid: 0,
      empId: "",
      peopleval: null,
      defaultroles: [],
      pages: 1,
      currentPage: 1,
      itemsPerPage: 5,
      currentData: [],
      itemsPerPageOptions: [5, 10, 20, 50, 100],
      onAddhide: false,
      onUpdatehide: false,
      valname: false,
      valempid: false,
      valmodule: false,
      valrole: false,
      selectedId: [],
      searchId: [],
      options: [],
      selectedName: [],
      searchName: [],
      selectedEmail: [],
      searchEmail: [],
      selectedRole: [],
      searchRole: [],
      SelectedDeptOptions: [],
      SearchDept: [],
      Depti: "",
      DeptOptions: [
        { label: "one", value: "one" },
        { label: "two", value: "two" },
        { label: "three", value: "three" },
      ],
      statusi: "",
      iscreatedby: [],
      searchcreate: [],
      // selectedRange: [],
      // statusOptions: [
      //   { label: "Active", value: "Active" },
      //   { label: "Inactive", value: "Inactive" },
      // ],
      datarecordsDup: [], //added by siva krishna
      DeleteId: "",
      DeleteStatus: false,
      validatemodulecheck: false,
      pagenumber: 1,
      selectedRange: null,
      isDataFound: true,
      IsUseradmin: [],
      role: '',
      EmPIDCheck: [],
      IsUserRepository: [],
      IsUserLead: [],
      IsUserTawazun: [],
      IsUserApex: [],
      ISUSerSouqAlMaal: [],
      IsUserShariaResearch: [],
      IsVertex: [],
      CheckRemover: [],
      RemoveID: [],
      RemovedRoleIds: [],


      ///siva buri////
      HelpdeskUrl: '',
      UAMUrl: '',
      ShariaRepoUrl: '',
      eLibraryUrl: '',
      ShariaResearchUrl: '',
      LeadXUrl: '',
      TawazunUrl: '',
      VertexUrl: '',
      ApexUrl: '',
      SouqAlMaalUrl: '',
      CheckEmployeeID: false,
      // accessToken: "",
      accessToken: "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiJhcGk6Ly9iMzZkMDhjOS1lNjViLTRhYTgtYWI3ZC0xZGM2YmIzZjAzN2EiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mZjQ5YzQzOC1jNDY5LTRjMTAtOTZmNi02MWY1NGRmNDFjOWIvIiwiaWF0IjoxNjk1NzIyMjAxLCJuYmYiOjE2OTU3MjIyMDEsImV4cCI6MTY5NTcyNjEwMSwiYWlvIjoiRTJGZ1lPaFFLTkVJVzd0ejBjbVQwYysrTEJXV0JRQT0iLCJhcHBpZCI6ImIzNmQwOGM5LWU2NWItNGFhOC1hYjdkLTFkYzZiYjNmMDM3YSIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2ZmNDljNDM4LWM0NjktNGMxMC05NmY2LTYxZjU0ZGY0MWM5Yi8iLCJvaWQiOiI4MmFhZDE1NS02YWQyLTRiMTUtYWVhOC0yM2I4NGRiOWZiODciLCJyaCI6IjAuQVVnQU9NUkpfMm5FRUV5VzltSDFUZlFjbThrSWJiTmI1cWhLcTMwZHhyc19BM3BJQUFBLiIsInJvbGVzIjpbImlkYW1fdXNlciJdLCJzdWIiOiI4MmFhZDE1NS02YWQyLTRiMTUtYWVhOC0yM2I4NGRiOWZiODciLCJ0aWQiOiJmZjQ5YzQzOC1jNDY5LTRjMTAtOTZmNi02MWY1NGRmNDFjOWIiLCJ1dGkiOiJiV3JLNE1rLWlFR3gtUWNCeE00VEFBIiwidmVyIjoiMS4wIn0.croHxFDOV3bTTcTZ0xjeCLMdVBaihTpFhF5vf0AXdd_2FJca7I0at1pZ60LLoCi42FSw3h-7EO1dbYY6nLT2FEQGBa5S4EywBXr2b3Gsw_fwiqKSlEoVpOSe00IJEG8r8-SP1MYxkgHe6f0Apt-jEz_c6l2kIegib0WRTFWnx8eW0lLYPcXUgYLGDi2uVJi6NFzrncHsZ5PwUlYrZHbg-ZaCEbfbKIkRvjBkGyDlqiwdwG1gpu9U4tb6EXr8E4_8ofL1s0YjyMhK6RgsOAUmwqI-Xhp4TWkoLtQXD53aT_V0WR7ormJgg3-_qbT64VRUthrQUpXEhvI4ycGK_4xcjA",

      isLoading: true,
      isHide: false,
    };
  }

  // public componentWillmount(): void {
  //   this.getAccessToken(this.props.email);
  // }

  public async componentDidMount() {

    await this.props.context.aadTokenProviderFactory.getTokenProvider().then((provider: any) => {
      provider.getToken("ed31b499-d6da-4250-8367-052fd153daf9").then((token: any) => {
        this.setState({ accessToken: token }, async () => {
          //console.log(token, "token form directry before getting Roles")


          //this.GetELibrarywebAppAPI();
        
          this.allUser();
          this.findUser();
          this.Application();
          this.Search();
          //this.Roles();
         
          this.GetApplicationModules();
        })
        //console.log(token, "token form directry");
      })


    })



    // await this.GetAuthentication();
    // this.getAccessToken(this.props.email);

    setTimeout(() => {
      this.setState({ isLoading: false });
    }, 3000);

    // const ID = this.state.empId
    // this.EmpIDDuplicate(ID);

    // this.renderpagination();




    //SIDE NAV CROSS & 3 LINE CSS
    $(".hamburger-left").on("click", function () {
      $(".hamburger").toggleClass("is-active");
    });

    //SIDE NAV OPEN AND CLOSE CSS
    $(".toggle-icon").on("click", function () {
      $(".wrapper").hasClass("toggled")
        ? ($(".wrapper").removeClass("toggled"),
          $(".sidebar-wrapper").unbind("hover"))
        : ($(".wrapper").addClass("toggled"),
          $(".sidebar-wrapper").on(
            "hover",
            function () {
              $(".wrapper").addClass("sidebar-hovered");
            },
            function () {
              $(".wrapper").removeClass("sidebar-hovered");
            }
          ));
      $(".wrapper").toggleClass("sidebar-hovered");
    });

    //SIDE NAV DROP DOWN CSS
    $(document).on("click", ".metismenu .menu-item.sub-menu a", function () {
      if ($(this).hasClass("active")) {
        $(this).toggleClass("active");
        $(this).closest("li").find(".submenu-dropdown").toggleClass("active");
      } else {
        $(".metismenu .menu-item a.menu-link").removeClass("active");
        $(".metismenu .menu-item .submenu-dropdown").removeClass("active");
        $(this).addClass("active");
        $(this).closest("li").find(".submenu-dropdown").addClass("active");
      }
    });
  }

  public async GetELibrarywebAppAPI() {
    try{
    let apiURL = `${this.context.pageContext.web.absoluteUrl}/_api/lists/GetByTitle('URL_Configuration')/items?$select=Title,WebAPiURL`;
    this.context.spHttpClient
      .get(apiURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          for (let i = 0; i < responseJSON.value.length; i++) {

          
              const check = responseJSON.value[i].Title
              //console.log(check, 'checkCondition')
              if (check === 'UserAccessManagementWebAPI') {
                Constants.var = responseJSON.value[i].WebAPiURL.Url;

                  this.allUser();
           this.Application();
          this.Search();
          this.Roles();
          this.findUser();
          this.GetApplicationModules();

              }
              else if(check === "UAMPowerBIUrl")
              {
                Constants.PowerBiReport = responseJSON.value[i].WebAPiURL.Url
                //console.log(Constants.PowerBiReport,"power bi report")
              }
            
          }
          //constants.shariaResearchWebapi =responseJSON.value[0].WebAPiURL.Url;
          //   this.setState({ allDataApi: responseJSON.value })
          //console.log(this.state.allDataApi, "web app apiurl");
        });
      });
    }
    catch(ex:any)
    {
      //console.log(ex);
    }
    // if(Webapi === "shariaResearchWebapi"){
    //   const ShariaResearchWebApi = this.state.allDataApi.map((obj:any)=>obj.WebAPiURL)
    // }
  }

  public async GetAuthentication() {
    let apiURL = `${this.props.siteUrl}/_api/lists/GetByTitle('ADAuthentication')/items?$select=Title,Value`;
    await this.props.context.spHttpClient.get(apiURL, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      response.json().then(async (responseJSON: any) => {
        debugger;
        let msalConfig: msal.Configuration = {
          auth: {
            clientId: `${responseJSON.value[0].Value}`,
            authority: `${responseJSON.value[1].Value}`
          },
          cache: {
            cacheLocation: "sessionStorage", // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
            storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge

          },
          system: {
            iframeHashTimeout: 10000,
            loggerOptions: {
              loggerCallback: (level, message, containsPii) => {
                if (containsPii) {
                  return;
                }
                switch (level) {
                  case msal.LogLevel.Error:
                    console.error(message);
                    return;
                  case msal.LogLevel.Info:
                    console.info(message);
                    return;
                  case msal.LogLevel.Verbose:
                    console.debug(message);
                    return;
                  case msal.LogLevel.Warning:
                    console.warn(message);
                    return;
                }
              },
            },
          },
        };

        // debugger;
        let msalInstance: msal.PublicClientApplication = new msal.PublicClientApplication(msalConfig);
        let currentAccount: msal.AccountInfo = null;
        let scopeSplit = responseJSON.value[2].Value && responseJSON.value[2].Value.split("'");
        let scope = [scopeSplit && scopeSplit.length > 0 && scopeSplit[1]];
        // console.log(scope, 'scope updated');
        let tokenrequest: any = {
          scopes: scope,
          //["Mail.Read"] ,//['api://c0f96cba-486b-4867-a2e2-e60ad996a3ed/hello'], ["Mail.Read"],
          //scopes:["User.Read"],
          account: currentAccount,
        };
        // debugger;
        // console.log(tokenrequest);
        // console.log(currentAccount);

        this.setState({ msalInstance: msalInstance, currentAccount: currentAccount, tokenrequest: tokenrequest }, async () => {
          await this.getAccessToken(this.props.context.pageContext.user.email);
        });

        // msalInstance = msalInstance1;
        // currentAccount = currentAccount1;
        // tokenrequest = tokenrequest1;


      });
    });
  }

  // public async GetPowerBiUrl() {
  //   let apiURL = `${this.props.siteUrl}/_api/lists/GetByTitle('ADAuthentication')/items?$select=Title,Value`;
  //   await this.props.context.spHttpClient.get(apiURL, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
  //     response.json().then(async (responseJSON: any) => {

  //     });
  //   });
  // }

  headerNav() {
    document.getElementById("hideHeaderNav").classList.toggle("d-none");
  }
  public Validation = () => {
    var isNameval: boolean = false;
    var isEmpidval: boolean = false;
    var isModuleval: boolean = false;
    var isRoleval: boolean = false;
    var ValidComplete: boolean = false;

    if (this.state.userName === "" || this.state.userName.length <= 0) {
      isNameval = true;
      this.setState({ valname: true });
    } else {
      this.setState({ valname: false });
    }

    if (this.state.empId.length <= 0) {
      isEmpidval = true;
      this.setState({ valempid: true });
    } else {
      this.setState({ valempid: false });
    }

    if (this.state.applicationSelectedValue.length <= 0) {
      isModuleval = true;
      this.setState({ valmodule: true });
    } else {
      this.setState({ valmodule: false });
    }

    if (this.state.applicationSelectedRoles.length <= 0) {
      isRoleval = true;
      this.setState({ valrole: true });
    } else {
      this.setState({ valrole: false });
    }

    if (isNameval === true || isEmpidval === true || isModuleval === true || isRoleval === true) {
      ValidComplete = false;
    } else {
      ValidComplete = true;
    }
    return ValidComplete;
  };
  // Validation = () => {
  //   var isNameval = this.state.userName === "" || this.state.userName.length <= 0;
  //   var isEmpidval = this.state.empId.length <= 0;
  //   var isModuleval = this.state.applicationSelectedValue.length <= 0;
  //   var isRoleval = this.state.applicationSelectedRoles.length <= 0;

  //   this.setState({ valname: isNameval });
  //   this.setState({ valempid: isEmpidval });
  //   this.setState({ valmodule: isModuleval });
  //   this.setState({ valrole: isRoleval });

  //   return !(isNameval || isEmpidval || isModuleval || isRoleval);
  // };
  public ValidationUpd = () => {
    var isModuleval: boolean = false;
    var isRoleval: boolean = false;
    var ValidComplete: boolean = false;
    if (this.state.applicationSelectedValue.length <= 0) {
      isModuleval = true;
      this.setState({ valmodule: true });
    } else {
      this.setState({ valmodule: false });
    }
    if (this.state.applicationSelectedRoles.length <= 0) {
      isRoleval = true;
      this.setState({ valrole: true });
    } else {
      this.setState({ valrole: false });
    }
    if (isRoleval === true || isModuleval === true) {
      ValidComplete = false;
    } else {
      ValidComplete = true;

    }
    return ValidComplete;
  };





  public Search() {
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}/User/UserList`,
      headers: { Authorization: `Bearer ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/UserList`
        config
      )
      .then((response) => {
        this.setState({ datarecords: [] }, () => {
          this.setState({ datarecords: response.data });
        });
      })
      .catch((error) => {
        //console.log(error);
       });
  }

  public allUser = () => {
    debugger;
    try {
      let config = {
        method: 'get',

        maxBodyLength: Infinity,

        url: `${Constants.var}/User/UserList`,
        headers: { Authorization: `Bearer  ${this.state.accessToken}` }
      }
      axios
        .request(
          config
        )
        .then((response) => {
          this.setState({ datarecords: [] }, () => {
           // console.log(response.data);
            this.setState({
              datarecords: response.data,
              datarecordsDup: response.data,
            }, () => { this.filterGrid(); });
          });
        }).catch((ex:any)=>{
          //console.log(ex);
        });
    } catch (exception) { }
  }

  public Application = () => {
    try {
      let config = {
        method: 'get',

        maxBodyLength: Infinity,

        url: `${Constants.var}User/Application`,
        headers: { Authorization: `Bearer   ${this.state.accessToken}` }
      }
      axios
        .request(
          // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/Application`
          config
        )
        .then((response) => {
          const values = response.data.map((option: any) => ({
            value: option.applicationId,
            label: option.applicationName,
          }));
          this.setState({ values });
        })
        .catch((error) => { 
          //console.log(error);
        });
    } catch (exception) { }
  };

  // public Add(e: any) {
  //   if (this.Validation()) {
  //     try {
  //       const { applicationSelectedRoles } = this.state;
  //       const defaultroles = applicationSelectedRoles.map(
  //         (option: any) => option.value
  //       );
  //       this.setState({ defaultroles: [] });
  //       var data = JSON.stringify({
  //         employeeId: this.state.empId,
  //         // id: this.state.empId,
  //         name: this.state.userName,
  //         roleIds: defaultroles,
  //         emailId: this.state.emailId,
  //         createdby: this.props.context.pageContext.user.email,
  //         Department: this.state.branch,
  //       });
  //       console.log(data);
  //       axios.post(
  //         // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/AddUser`,
  //         `${Constants.var}User/AddUser`,
  //         {
  //           method: "POST",
  //           headers: {
  //             "Content-Type": "application/json",

  //           },
  //           data: data,
  //         }
  //       )
  //         .then((response) => {
  //           console.log(JSON.stringify(response.data));
  //           document.getElementById("Close-Modal").click();
  //         })
  //         .then(() => this.Search())
  //         .then(() =>
  //           this.setState({
  //             employeeID: "",
  //             empId: "",
  //             userName: "",
  //             emailId: "",
  //             branch: "",
  //             applicationSelectedRoles: [],
  //             applicationSelectedValue: [],
  //             roles: [],
  //             isActive: true,
  //             peopleval: null,
  //             onAddhide: false,
  //             onUpdatehide: false,
  //             valname: false,
  //             valempid: false,
  //             valrole: false,
  //             valmodule: false,
  //           })
  //         );
  //       swal.fire({
  //         title: "Success!",
  //         text: "User Added Successfully!",
  //         icon: "success",
  //         showCancelButton: false,
  //       })
  //         .then(() => this.allUser())
  //         .catch((error) => {
  //           console.log(error);
  //         });
  //     } catch (exception) { }
  //   }
  // }

  public Add(e: any) {
    if (this.Validation()) {
      try {
        const { applicationSelectedRoles } = this.state;
        const defaultroles = applicationSelectedRoles.map(
          (option: any) => option.value
        );
        this.setState({ defaultroles: [] });
        const EmployeeID = parseInt(this.state.empId)
        const duplicateemployeeId = this.state.datarecords.map((obj: any) => obj.employeeId).join(',');
        // console.log(duplicateemployeeId, 'duplicates')

        if (duplicateemployeeId.includes(EmployeeID)) {
          swal.fire("Info", "Employee Id you provided already exist. Please enter another one.", 'info');
          return;
        }
        // let config = {
        //   headers: { Authorization: `Bearer ${Constants.AccessToken}` }
        // }


        var data = JSON.stringify({
          employeeId: this.state.empId,
          // id: this.state.empId,
          name: this.state.userName,
          roleIds: defaultroles,
          emailId: this.state.emailId,
          createdby: this.props.context.pageContext.user.email,
          Department: this.state.branch,
        });
        // console.log(data);
        axios(
          // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/AddUser`,
          `${Constants.var}User/AddUser`,
          {
            method: "post",
            headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer  ${this.state.accessToken}`
            },
            data: data,
          }
        )
          .then((response) => {
            // console.log(JSON.stringify(response.data));
            document.getElementById("Close-Modal").click();
          })
          .then(() => this.Search())
          .then(() =>
            this.setState({
              employeeID: "",
              empId: "",
              userName: "",
              emailId: "",
              branch: "",
              applicationSelectedRoles: [],
              applicationSelectedValue: [],
              roles: [],
              isActive: true,
              peopleval: null,
              onAddhide: false,
              onUpdatehide: false,
              valname: false,
              valempid: false,
              valrole: false,
              valmodule: false,
            })
          );
        swal.fire({
          title: "Success!",
          text: "User Added Successfully!",
          icon: "success",
          showCancelButton: false,
        })
          .then(() => this.allUser())
          .catch((error) => {
            console.log(error);
          });
      } catch (exception) { }
    }
  }

  public Roles = () => {
    try {
      let config = {
        method: 'get',

        maxBodyLength: Infinity,

        url: `${Constants.var}User/RolesByApplication?applicationId=${this.state.applicationSelectedValue.value}`,
        headers: { Authorization: `Bearer  ${this.state.accessToken}` }
      }

      const { applicationSelectedValue } = this.state;
      if (applicationSelectedValue) {
        axios
          .request(
            config
          )
          .then((response) => {
            const roles = response.data.map((option: any) => ({
              value: option.roleId,
              label: option.name,
            }));
            this.setState({ roles },
               () => this.modifydata(roles,"setapplicationroles")
            );
          })
          .catch((error) => {
            console.error("Error fetching roles:", error);
            this.setState({ roles: [] });
          });
      } else {
        this.setState({ roles: [], applicationSelectedRoles: [] });
      }
    } catch (exception) {
      console.error("Exception occurred:", exception);
      this.setState({ roles: [] });
    }
  };

  public getdatabyid = (id: any) => {
    let config = {
      method: 'get',
      maxBodyLength: Infinity,
      url: `${Constants.var}User/getUserByUserId?UserId=${id}`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(config)
      .then((response) => {
        if (response && response.data && response.data.length > 0) {
          const applicationSelectedRoles = response.data[0].roles.map(
            (option: any) => ({
              value: option.roleId,
              label: option.name,
            })
          );
          const applicationSelectedValue = {
            value: response.data[0].roles[0].applicationId,
            label: response.data[0].roles[0].applicationName,
          };
          this.setState(
            {
              applicationSelectedRoles,
              applicationSelectedValue,
              employeeID: response.data[0].employeeId,
              emailId: response.data[0].emailId,
              userName: response.data[0].userName,
              branch: response.data[0].department,
              Userid: response.data[0].id,
              onAddhide: false,
              onUpdatehide: true,
            },
            () => {
              this.Roles();
              this.modifydata(applicationSelectedRoles,"fromgetbyid");

            }
          );
         


        } else {

          this.setState({ roles: [], applicationSelectedRoles: [] });
        }
      })
      .catch((error) => {
        console.error("Error fetching user data:", error);
        this.setState({ roles: [], applicationSelectedRoles: [] });
      });


    // this.onchnagemoduleValidation()
  };







  public update(e: any) {
    if (this.ValidationUpd()) {
      const { applicationSelectedRoles } = this.state;

      const defaultroles = applicationSelectedRoles.map(
        (option: any) => option.value
      );
      //const defaultroles = applicationSelectedRoles.filter(())
      this.setState({ defaultroles: [] });

      // const a = defaultroles.find((obj: any) => obj == this.state.RemoveID)

      // const b = this.state.RemoveID.splice([a])
      // // this.setState({RemoveID:b})
      // console.log(b, 'Removal')

      // console.log(this.state.valrole);
      var data = JSON.stringify({
        id: this.state.Userid,
        // employeeId: this.state.employeeID,
        // name: this.state.userName
        RemovedRoleIds: this.state.RemoveID,
        roleIds: defaultroles,
        // emailId: this.state.emailId,
        createdby: this.props.context.pageContext.user.email,
      });
      axios(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/UpdateUserRole`,
        `${Constants.var}User/UpdateUserRole`,
        {
          method: "put",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${this.state.accessToken}`
          },
          data: data,
        }
      )
        .then((response) => {
          document.getElementById("Close-Modal").click();
        })
        .then(() => this.Search())
        .then(() =>
          this.setState({
            employeeID: "",
            userName: "",
            emailId: "",
            branch: "",
            applicationSelectedRoles: [],
            applicationSelectedValue: [],
            roles: [],
            isActive: true,
            peopleval: null,
            onAddhide: false,
            onUpdatehide: false,
            valname: false,
            valempid: false,
            valrole: false,
            valmodule: false,
          })
        );
      swal.fire({
        title: "Success!",
        text: "User Update Successfully!",
        icon: "success",
        showCancelButton: false,
      })
        .then(() => this.allUser())
        .catch((error) => {
          console.error("Error updating user:", error);
          swal.fire({
            title: "Error",
            text: "Failed to update user. Please try again later.",
            icon: "error",
            showCancelButton: false,
          });
        });
    }
  }

  // public handleroles = (applicationSelectedRoles: any) => {

  //   this.setState({ applicationSelectedRoles }, () => {
  //     this.updateRemoveID();
  //   });
  // };

  // private updateRemoveID = () => {
  //   const { applicationSelectedRoles } = this.state;
  //   const defaultroles = applicationSelectedRoles.map((option: any) => option.value);

  //   // Assuming RemoveID is an array of role IDs to be removed
  //   // Update RemoveID based on defaultroles
  //   const updatedRemoveID = this.state.RemoveID.filter((id: any) => defaultroles.includes(id));

  //   // Update RemoveID in state
  //   this.setState({ RemoveID: updatedRemoveID }, () => {
  //     console.log('Updated RemoveID:', this.state.RemoveID);
  //   });
  // };




  // public update = (e: any) => {
  //   if (this.ValidationUpd()) {
  //     const { applicationSelectedRoles, RemovedRoleIds } = this.state;
  //     const defaultroles = applicationSelectedRoles.map((option: any) => option.value);

  //     const data = JSON.stringify({
  //       id: this.state.Userid,
  //       RemovedRoleIds: RemovedRoleIds,
  //       roleIds: defaultroles,
  //       createdby: this.props.context.pageContext.user.email,
  //     });

  //     axios(
  //       `${Constants.var}User/UpdateUserRole`,
  //       {
  //         method: "put",
  //         headers: {
  //           "Content-Type": "application/json",
  //           Authorization: `Bearer ${this.state.accessToken}`,
  //         },
  //         data: data,
  //       }
  //     )
  //       .then(() => {
  //         document.getElementById("Close-Modal").click();
  //         return this.Search();
  //       })
  //       .then(() => {
  //         this.setState({
  //           employeeID: "",
  //           userName: "",
  //           emailId: "",
  //           branch: "",
  //           applicationSelectedRoles: [],
  //           applicationSelectedValue: [],
  //           roles: [],
  //           isActive: true,
  //           peopleval: null,
  //           onAddhide: false,
  //           onUpdatehide: false,
  //           valname: false,
  //           valempid: false,
  //           valrole: false,
  //           valmodule: false,
  //         });
  //       })
  //       .then(() => {
  //         swal.fire({
  //           title: "Success!",
  //           text: "User Update Successfully!",
  //           icon: "success",
  //           showCancelButton: false,
  //         });
  //         return this.allUser();
  //       })
  //       .catch((error) => {
  //         console.error("Error updating user:", error);
  //         swal.fire({
  //           title: "Error",
  //           text: "Failed to update user. Please try again later.",
  //           icon: "error",
  //           showCancelButton: false,
  //         });
  //       });
  //   }
  // };










  public toDateString(date: string) {
    const buffer = new Date(date);
    return `${buffer.getDate()}/${buffer.getMonth()}/${buffer.getFullYear()}`;
  }

  public async department(EmailId: any) {
    const accountName = encodeURIComponent(`i:0#.f|membership|${EmailId}`);
    const userProfileUrl = `${this.props.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName='Department')?@v='${accountName}'`;
    //  alert(userProfileUrl) 
    const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(userProfileUrl, SPHttpClient.configurations.v1);
    const userProfile = await response.json();
    const department
      = userProfile['value'];
    // console.log(department);
    // alert  (department),
    this.setState({ branch: department });
    // console.log(this.state.branch);
    // alert(this.state.branch)
  }

  public onChangePeoplePickerItems = (e: any, branch: any) => {
    var context = this;
    var empId = e[0].id;
    var name = e[0].text;
    var EmailId = e[0].secondaryText;
    var isAvailable = false;


    for (let i = 0; i < this.state.datarecords.length; i++) {
      if (EmailId == this.state.datarecords[i].emailId) {
        isAvailable = true;
      }
    }
    // for (let i = 0; i < this.state.datarecords.length; i++) {
    //   if (empId == this.state.datarecords[i].id && EmailId == this.state.datarecords[i].emailId) {
    //     isAvailable = true;
    //   }
    // }
    if (isAvailable == true) {
      context.setState({
        employeeID: "",
        branch: "",
        userName: "",
        emailId: "",
        isActive: false,

      });
    } else {
      context.setState({
        employeeID: empId,
        userName: name,
        emailId: EmailId,
        branch,
        isActive: true,

      });
    }
    this.forceUpdate();
    this.department(EmailId);
    // this.Validation();
  };

  handleChangeId = (e: any) => {
    //console.log(e);
    this.setState({ selectedId: e }, () => {
      this.filterGrid(); //added by siva krishna
    });
  };

  handleInputChangeId = (e: any) => {
    const a = e.replace(/"/g, "");
    this.setState({ searchId: e });
   // console.log(a, "IdSearch");
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}User/SearchUserByEmployeeId?query=${a}`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/SearchUserByEmployeeId?query=${a}`
        config
      )
      .then((response) => {
        //console.log(response.data);
        const options = response.data.map((obj: any) => ({
          value: obj,
          label: obj,
        }));
        this.setState({ options });
      })
      .catch((error) => {
        //console.log(error);
      });
  };

  handleChangeName = (e: any) => {
    //console.log(e);
    this.setState({ selectedName: e }, () => {
      this.filterGrid(); //added by siva krishna
    });
  };

  handleInputChangeName = (e: any) => {
    const a = e.replace(/"/g, "");
    this.setState({ searchName: e });
   // console.log(a, "NameSearch");
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}User/SearchUserByName?query=${a}`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/SearchUserByName?query=${a}`
        config
      )
      .then((response) => {
       // console.log(response.data);
        const options = response.data.map((obj: any) => ({
          value: obj,
          label: obj,
        }));
        this.setState({ options });
      })
      .catch((error) => {
        //console.log(error);
      });
  };

  handleChangeEmail = (e: any) => {
    //console.log(e);
    this.setState({ selectedEmail: e }, () => {
      this.filterGrid(); //added by siva krishna
    });
  };

  handleInputChangeEmail = (e: any) => {
    const a = e.replace(/"/g, "");
    this.setState({ searchEmail: e });
    //console.log(a, "EmailSearch");
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}User/SearchUserByEmail?query=${a}`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/SearchUserByEmail?query=${a}`
        config
      )
      .then((response) => {
        //console.log(response.data);
        const options = response.data.map((obj: any) => ({
          value: obj,
          label: obj,
        }));
        this.setState({ options });
      })
      .catch((error) => {
        //console.log(error);
      });
  };

  handleChangeRole = (e: any) => {
    //console.log(e);
    this.setState({ selectedRole: e }, () => {
      this.filterGrid();
    });
  };

  handleInputChangeRole = (e: any) => {
    const a = e.replace(/"/g, "");
    this.setState({ searchRole: e });
    //console.log(a, "RoleSearch");
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}User/SearchRoles?query=${a}`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/SearchRoles?query=${a}`
        config
      )
      .then((response) => {
        //console.log(response.data);
        const options = response.data.map((obj: any) => ({
          value: obj,
          label: obj,
        }));
        this.setState({ options });
      })
      .catch((error) => {
        //console.log(error);
      });
  };

  handleChangecreate = (e: any) => {
    //console.log(e);
    this.setState({ iscreatedby: e }, () => {
      this.filterGrid();
    });
  };

  handleInputChangecreate = (e: any) => {
    const a = e.replace(/"/g, "");
    this.setState({ searchcreate: e });
   // console.log(a, "createSearch");
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}User/SearchCreatedBy?query=${a}`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/SearchCreatedBy?query=${a}`
        config
      )
      .then((response) => {
        //console.log(response.data);
        const options = response.data.map((obj: any) => ({
          value: obj,
          label: obj,
        }));
        this.setState({ options });
      })
      .catch((error) => {
        //console.log(error);
      });
  };

  handleChangeDept = (e: any) => {
    //console.log(e);
    this.setState({ SelectedDeptOptions: e }, () => {
      this.filterGrid();
    });
  };

  handleInputChangeDept = (e: any) => {
    //const a = e.replace(/"/g, "");
    this.setState({ SearchDept: e });
    //console.log(a, "DeptSearch");
    let config = {
      method: 'get',

      maxBodyLength: Infinity,

      url: `${Constants.var}User/GetDepartment?department`,
      headers: { Authorization: `Bearer  ${this.state.accessToken}` }
    }
    axios
      .request(
        // `https://almasdar-useraccessmanagement.azurewebsites.net/api/User/SearchCreatedBy?query=${a}`
        config
      )
      .then((response) => {
       // console.log(response.data);
        var options = response.data.map((obj: any) => ({
          value: obj.department,
          label: obj.department,
        }));
        //var options =[{ label: response.data[0].department, value : response.data[0].department}]
        this.setState({ DeptOptions: options });
        // alert(options);
        //console.log(options)
      })
      .catch((error) => {
        //console.log(error);
      });
  };
  handleDateRangeChange = (dateRange: any) => {
    this.setState({ selectedRange: dateRange }, () => {
      this.filterGrid()
    })
  };


  // public async filterGrid() {
  //   // try {
  //   //   let filData: any = [];
  //   try {
  //     const {
  //       selectedId,
  //       selectedName,
  //       selectedEmail,
  //       selectedRole,
  //       iscreatedby,
  //       selectedRange,
  //       SelectedDeptOptions,
  //       datarecords,
  //     } = this.state;
  //     let filData: any = [];
  //     // Filtering selected employee IDs
  //     this.state.selectedId &&
  //       this.state.selectedId.length > 0 &&
  //       this.state.selectedId.map((eOpt: any) => {
  //         this.state.datarecords &&
  //           this.state.datarecords.length > 0 &&
  //           this.state.datarecords.map((opt: any) => {
  //             if (opt.employeeId === eOpt.value) {
  //               filData.push(opt);
  //             }
  //           });
  //       });
  //     // Filtering selected names
  //     this.state.selectedName &&
  //       this.state.selectedName.length > 0 &&
  //       this.state.selectedName.map((eOpt: any) => {
  //         this.state.datarecords &&
  //           this.state.datarecords.length > 0 &&
  //           this.state.datarecords.map((opt: any) => {
  //             if (opt.userName === eOpt.value) {
  //               filData.push(opt);
  //             }
  //           });
  //       });
  //     // Filtering selected emails
  //     this.state.selectedEmail &&
  //       this.state.selectedEmail.length > 0 &&
  //       this.state.selectedEmail.map((eOpt: any) => {
  //         this.state.datarecords &&
  //           this.state.datarecords.length > 0 &&
  //           this.state.datarecords.map((opt: any) => {
  //             if (opt.emailId === eOpt.value) {
  //               filData.push(opt);
  //             }
  //           });
  //       });
  //     // Filtering selected roles
  //     this.state.selectedRole &&
  //       this.state.selectedRole.length > 0 &&
  //       this.state.selectedRole.map((eOpt: any) => {
  //         this.state.datarecords &&
  //           this.state.datarecords.length > 0 &&
  //           this.state.datarecords.map((opt: any) => {
  //             opt.roles &&
  //               opt.roles.length > 0 &&
  //               opt.roles.map((role: any) => {
  //                 if (role.name === eOpt.value) {
  //                   filData.push(opt);
  //                 }
  //               });
  //           });
  //       });
  //     // Filtering selected created by
  //     this.state.iscreatedby &&
  //       this.state.iscreatedby.length > 0 &&
  //       this.state.iscreatedby.map((eOpt: any) => {
  //         this.state.datarecords &&
  //           this.state.datarecords.length > 0 &&
  //           this.state.datarecords.map((opt: any) => {
  //             if (opt.createdby === eOpt.value) {
  //               filData.push(opt);
  //             }
  //           });
  //       });
  //     //fitering date
  //     // if (this.state.selectedRange && this.state.selectedRange.length === 2) {
  //     //   const [startDate, endDate] = this.state.selectedRange;
  //     //   console.log('Start Date:', startDate.format('YYYY-MM-DD'));
  //     //   console.log('End Date:', endDate.format('YYYY-MM-DD'));
  //     //   const filteredDateRange = this.state.datarecords.filter((opt: any) => {
  //     //     const createdDate = moment(opt.createdDate);
  //     //     return createdDate.isBetween(startDate, endDate, 'day', '[]');
  //     //   });
  //     //   filData.push(...filteredDateRange);
  //     //   await this.setState({
  //     //     datarecords: filData
  //     //   })
  //     // }

  //     if (selectedRange && selectedRange.length === 2) {
  //       const [startDate, endDate] = selectedRange;
  //       filData = filData.filter((opt: any) => {
  //         const createdDate = moment(opt.createdDate);
  //         return createdDate.isBetween(startDate, endDate, 'day', '[]');
  //       });
  //     }
  //     // Filtering selected department options
  //     this.state.SelectedDeptOptions &&
  //       this.state.SelectedDeptOptions.length > 0 &&
  //       this.state.SelectedDeptOptions.map((eOpt: any) => {
  //         this.state.datarecords &&
  //           this.state.datarecords.length > 0 &&
  //           this.state.datarecords.map((opt: any) => {
  //             if (opt.department === eOpt.value) {
  //               filData.push(opt);
  //             }
  //           });
  //       });
  //     // if (
  //     //   this.state.selectedId &&
  //     //   this.state.selectedId.length === 0 &&
  //     //   this.state.selectedName &&
  //     //   this.state.selectedName.length === 0 &&
  //     //   this.state.selectedEmail &&
  //     //   this.state.selectedEmail.length === 0 &&
  //     //   this.state.selectedRole &&
  //     //   this.state.selectedRole.length === 0 &&
  //     //   this.state.iscreatedby &&
  //     //   this.state.iscreatedby.length === 0 &&
  //     //   this.state.SelectedDeptOptions &&
  //     //   this.state.SelectedDeptOptions.length === 0 &&
  //     //   (!this.state.selectedRange || this.state.selectedRange.length !== 2)
  //     // ) {
  //     //   this.setState({ datarecordsDup: this.state.datarecords });
  //     //   return;
  //     // }
  //     if (
  //       selectedId.length === 0 &&
  //       selectedName.length === 0 &&
  //       selectedEmail.length === 0 &&
  //       selectedRole.length === 0 &&
  //       iscreatedby.length === 0 &&
  //       SelectedDeptOptions.length === 0 &&
  //       (!selectedRange || selectedRange.length !== 2)
  //     ) {
  //       this.setState({
  //         datarecordsDup: datarecords,
  //         currentPage: 1,
  //       });
  //       return;
  //     }
  //     filData = filData.filter(
  //       (ele: any, ind: any) =>
  //         ind ===
  //         filData.findIndex((elem: any) => elem.employeeId === ele.employeeId)
  //     );
  //     if (filData.length === 0) {
  //       this.setState({ datarecordsDup: this.state.datarecords });
  //     } else {
  //       this.setState({ datarecordsDup: filData });
  //     }
  //     await this.setState({
  //       datarecordsDup: filData,
  //       currentPage: 1,
  //     });
  //   } catch (ex) {
  //     console.log(ex.message);
  //   }
  // }
  public async filterGrid() {
    try {
      const {
        selectedId,
        selectedName,
        selectedEmail,
        selectedRole,
        iscreatedby,
        selectedRange,
        SelectedDeptOptions,
        datarecords,
      } = this.state;
      if (
        selectedId.length === 0 &&
        selectedName.length === 0 &&
        selectedEmail.length === 0 &&
        selectedRole.length === 0 &&
        iscreatedby.length === 0 &&
        SelectedDeptOptions.length === 0 &&
        (!selectedRange || selectedRange.length !== 2)
      ) {
        this.setState({
          datarecordsDup: datarecords,
          currentPage: 1,
        });
        return;
      }
      let filData = datarecords;
      // Filtering selected employee IDs
      if (selectedId && selectedId.length > 0) {
        filData = filData.filter((opt: any) => {
          return selectedId.some((eOpt: any) => opt.employeeId === eOpt.value);
        });
      }
      // Filtering selected names
      if (selectedName && selectedName.length > 0) {
        filData = filData.filter((opt: any) => {
          return selectedName.some((eOpt: any) => opt.userName === eOpt.value);
        });
      }
      // Filtering selected emails
      if (selectedEmail && selectedEmail.length > 0) {
        filData = filData.filter((opt: any) => {
          return selectedEmail.some((eOpt: any) => opt.emailId === eOpt.value);
        });
      }
      // Filtering selected roles
      if (selectedRole && selectedRole.length > 0) {
        filData = filData.filter((opt: any) => {
          return selectedRole.some((eOpt: any) => {
            return opt.roles && opt.roles.some((role: any) => role.name === eOpt.value);
          });
        });
      }
      // Filtering selected created by
      if (iscreatedby && iscreatedby.length > 0) {
        filData = filData.filter((opt: any) => {
          return iscreatedby.some((eOpt: any) => opt.createdby === eOpt.value);
        });
      }
      // Filter by date range
      if (selectedRange && selectedRange.length === 2) {
        const [startDate, endDate] = selectedRange;
        filData = filData.filter((opt: any) => {
          const createdDate = moment(opt.createdDate);
          return createdDate.isBetween(startDate, endDate, 'day', '[]');
        });
      }
      // Filtering selected department options
      if (SelectedDeptOptions && SelectedDeptOptions.length > 0) {
        filData = filData.filter((opt: any) => {
          return SelectedDeptOptions.some((eOpt: any) => opt.department === eOpt.value);
        });
      }
      await this.setState({
        datarecordsDup: filData,
        currentPage: 1,
      });
    } catch (ex) {
      //console.log(ex.message);
    }
  }

 public async findUser() {
    debugger;
    let config = {
      headers: { Authorization: `Bearer ${this.state.accessToken}` }
    }
    axios
      .get(`${Constants.var}User/GetDetailsByEmailId?EmailId=` + this.props.context.pageContext.user.email, config)
      .then(response => {
        if (response.data && response.data.length > 0) {
          const roles1 = response.data[0].roles.map((role: { name: any }) => role.name);
          //console.log(roles1);

          // const roles = response.data[0].roles.map((role: { name: any; applicationName: any; }) => ({
          //   name: role.name,
          //   applicationName: role.applicationName
          // }));
          // console.log(roles, 'get all roles');

          // for (let i = 0; i <= roles.length; i++) {
          //   const Name = roles[i].name;
          //   const ApplicationNAme = roles[i].applicationName;
          //   if (Name === "Auditor" || Name === "ISCG Staff" || Name === "ISCG Staff Admin" && ApplicationNAme ===
          //     "Sharia Repository") {
          //     console.log(this.state.IsUserRepository, 'Shariya Repostory role');
          //   }
          //   else if (Name === "Helpdesk Manager" || Name === "ISCG Staff" || Name === "Master Data Admin - Sharia Research" && ApplicationNAme ===
          //     "Sharia Research") {
          //     console.log(this.state.IsUserShariaResearch, 'Shariya Reasearch role')
          //   }
          //   else {
          //     console.log('You  Dont have access to this page')
          //   }

          // }



          // const ChecknewValidate = roles.join('').map((obj:any)=>({
          //   label:obj.name,
          //   value:obj.applicationName
          // }))
          // console.log(ChecknewValidate,'new roles')

          this.setState({

            IsUseradmin: roles1,
            IsUserRepository: roles1,
            IsUserLead: roles1,
            IsUserTawazun: roles1,
            IsUserApex: roles1,
            ISUSerSouqAlMaal: roles1,
            IsUserShariaResearch: roles1,
            IsVertex: roles1

          });
        }
      })
      .catch(error => {
        console.error(error);
      });
  }





  public render(): React.ReactElement<IUserAccessManagementProps> {
    // const startIndex = (this.state.currentPage - 1) * this.state.itemsPerPage;
    // const endIndex = startIndex + this.state.itemsPerPage;
    // const currentData = this.state.datarecordsDup.slice(startIndex, endIndex);
    // console.log(this.state.datarecords, 'data')
    // console.log(this.state.userName, 'username');
    const { currentPage, itemsPerPage, datarecordsDup } = this.state;
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentData = datarecordsDup.slice(startIndex, endIndex);
    //console.log(this.state)


    return (
      <section>
        <body>
          {this.state.isLoading && (
            <div>
              <p id="pre-loader"></p>
            </div>
          )}
          <div className="wrapper help-desk-Homepage toggled">

            <header className="row flex-wrap align-items-center justify-content-center justify-content-md-center border-bottom bg-primary-8 h-55 px-5 w-100 m-0 pe-0 position-fixed z-index-9 ps-0">
              <div id="flipkart-navbar" className="UAM-header">
                <div className="container-fluid mt-1">
                  <div className="d-flex justify-content-between">
 
                    <div className="sidebar-header border-0">
                      <div className="toggle-icon ms-1 d-flex align-items-center">
                        {" "}
                        <button
                          type="button"
                          className="hamburger hamburger-left "
                        >
                          <span className="hamburger-box">
                            <span className="hamburger-inner"></span>
                          </span>
                        </button>
                      </div> 
                    </div>


                    <div className="flipkart-navbar-search smallsearch col-sm-8 col-xl-10 mt-8">
                      <ul
                        className="pull-right mb-0 ps-3 ms-6 pt-2 d-lg-block d-none mt-2"
                        id="hideHeaderNav"
                      >
                        <li className="upper-links default-active active"><a className="links mb-0 mt-1" href="#">HOME</a>
                        </li>
                        {Constants.PowerBiReport!=""&& <li className="upper-links"><a className="links mb-0 mt-1" onClick={()=>this.openpowerBI()}>REPORTS</a>
                        </li> }
                       
                        {/* <li className="upper-links"><a className="links fontsize-16 mb-0 mt-1" href="#/createrequest/NewRequest">Create Request</a>
                    </li> */}


                      </ul>
                    </div>
                    <div className='logo-right justify-content-between'>
                      <span className="largenav ms-auto">
                        <img
                          src={
                            this.props.context.pageContext.web.absoluteUrl +
                            "/siteAssets/assets/images/images-folder/MicrosoftTeams-image5.png"
                          } className='h-30'
                        />
                      </span>
                    </div>

                    <h2 className="smallnav ">
                      <div className=" menu mt-3">
                        <div className="d-flex justify-content-end">
                          <div>
                            <div>
                              <i
                                className="bi bi-three-dots-vertical"
                                onClick={() => {
                                  this.headerNav();
                                }}
                              ></i>
                            </div>
                          </div>
                        </div>
                      </div>
                    </h2>

                  </div>
                </div>
              </div>
            </header>

            

            <aside className="sidebar-wrapper border-0" data-simplebar="init">
              <div className="simplebar-wrapper bg-white shadow-lg">
                <div className="simplebar-height-auto-observer-wrapper">
                  <div className="simplebar-height-auto-observer">
                  </div>
                </div> 
                <div className="simplebar-mask"> 

                  <div className="simplebar-offset">
                    <div className="simplebar-content-wrapper">
                      <div className="simplebar-content"> 
                        <ul className="metismenu h-100" id="menu">
                          <li className="menu-item sub-menu pt-2">
                            <a
                              href="javascript:;"
                              className="menu-link has-arrow p-1 text-decoration-none ms-1"
                            >
                              <div className="parent-icon shadow p-1 shadow rounded">
                                <img
                                  src={
                                    this.props.siteUrl +
                                    "/siteAssets/assets/images/images-folder/fingerprint.png"
                                  }
                                  className="h-30 w-31 "
                                />
                              </div>
                              <div className="menu-title bliss2M fontsize-18">
                                Al Masdar
                              </div>
                            </a>
                            <ul className="submenu-dropdown ms-3">
                              <li className="menu-item">
                                <a
                                  // href={this.state.HelpdeskUrl} target="_blank" data-interception="off"
                                  onClick={() => this.openHelpdesk()}
                                  className="menu-link ps-1 text-decoration-none"
                                >
                                  <div className="p-1  rounded">
                                    <img
                                      src={
                                        this.props.siteUrl +
                                        "/siteAssets/assets/images/images-folder/Helpdesk-sidenav-Icon.png"
                                      }
                                      className="submenu-iconz" />
                                  </div>
                                  <div className="menu-title bliss2M fontsize-16 ms-4">
                                    Helpdesk
                                  </div>
                                </a>
                              </li>
                              {this.state.IsUserShariaResearch.includes("Helpdesk Manager-Sharia Research") || this.state.IsUserShariaResearch.includes("ISCG Staff - Sharia Research") || this.state.IsUserShariaResearch.includes("Master Data Admin - Sharia Research") ?
                                <li className="menu-item">
                                  <a
                                    // href="https://adibcloud.sharepoint.com/sites/Shariaresearch-uat" target="_blank" data-interception="off"
                                    onClick={() => this.openShariaReasearch()}
                                    className="menu-link ps-1 text-decoration-none ">
                                    <div className="p-1  rounded">
                                      <img
                                        src={
                                          this.props.siteUrl +
                                          "/siteAssets/assets/images/images-folder/sharia-research.png"
                                        }
                                        className="submenu-iconz" />
                                    </div>
                                    <div className="menu-title bliss2M fontsize-16 ms-4">
                                      Sharia Research
                                    </div>
                                  </a>
                                </li>
                                : null}

                              {this.state.IsUserTawazun.includes("Maker - Tawazun") || this.state.IsUserTawazun.includes("Checker - Tawazun") || this.state.IsUserTawazun.includes("Master Data Admin - Tawazun") ?

                                <li className="menu-item">
                                  <a
                                    // href="index.html" target="_blank" data-interception="off"
                                    onClick={() => this.openTawazun()}
                                    className="menu-link ps-2 text-decoration-none"
                                  >
                                    <div className="p-1 rounded">
                                      <img
                                        src={
                                          this.props.siteUrl +
                                          "/siteAssets/assets/images/images-folder/Tawazun-Sidenav.png"
                                        }
                                        className="submenu-iconz"
                                      />
                                    </div>
                                    <div className="menu-title bliss2M fontsize-16 ms-4">
                                      Tawazun
                                    </div>
                                  </a>
                                </li>
                                : null}

                              {this.state.IsUserApex.includes("Maker - Apex") || this.state.IsUserApex.includes("Checker - Apex") || this.state.IsUserApex.includes("Master Data Admin - Apex") ?

                                <li className="menu-item">
                                  <a
                                    // href="index.html" target="_blank" data-interception="off"
                                    onClick={() => this.openapex()}
                                    className="menu-link ps-2 text-decoration-none"
                                  >
                                    <div className="p-1  rounded">
                                      <img
                                        src={
                                          this.props.siteUrl +
                                          "/siteAssets/assets/images/images-folder/APEX-sidenav.png"
                                        }
                                        className="submenu-iconz"
                                      />
                                    </div>
                                    <div className="menu-title bliss2M fontsize-16 ms-4 ps-2">
                                      Apex
                                    </div>
                                  </a>
                                </li>
                                : null}

                              {this.state.IsVertex.includes("Maker - SAM") || this.state.IsVertex.includes("Checker - SAM") || this.state.IsVertex.includes("Master Data Admin - SAM") ?

                                <li className="menu-item">
                                  <a
                                    // href="index.html" target="_blank" data-interception="off"
                                    onClick={() => this.openvertex()}
                                    className="menu-link ps-2 text-decoration-none"
                                  >
                                    <div className="p-1 rounded">
                                      <img
                                        src={
                                          this.props.siteUrl +
                                          "/siteAssets/assets/images/images-folder/VertexSideNav.png"
                                        }
                                        className="submenu-iconz"
                                      />
                                    </div>
                                    <div className="menu-title bliss2M fontsize-16 ms-4">
                                      Vertex
                                    </div>
                                  </a>
                                </li>
                                : null}


                            </ul>

                          </li>
                          <li className="menu-item sub-menu">
                            <a
                              href="javascript:;"
                              className="menu-link has-arrow p-1 ms-1"
                              data-bs-toggle="collapse"
                              data-bs-target="#second-collapse"
                              aria-expanded="false"
                            >
                              <div className="parent-icon shadow p-1 shadow rounded">
                                <img
                                  src={
                                    this.props.siteUrl +
                                    "/siteAssets/assets/images/images-folder/AlifSidenav.png"
                                  }
                                  className=" w-31 aliff"
                                />
                              </div>

                              <div className="menu-title bliss2M fontsize-18">
                                Alif
                              </div>
                            </a>

                            <ul className="submenu-dropdown ms-2">
                              <li className="menu-item">
                                <a
                                  // href="https://adibcloud.sharepoint.com/sites/elibrary-uat" target="_blank" data-interception="off"
                                  onClick={() => this.openElibrary()}
                                  className="menu-link ps-2 text-decoration-none"
                                >
                                  <div className="p-1 rounded">
                                    <img
                                      src={
                                        this.props.siteUrl +
                                        "/siteAssets/assets/images/images-folder/ELibsideNav.png"
                                      }
                                      className="submenu-iconz"
                                    />
                                  </div>
                                  <div className="menu-title bliss2M fontsize-16 ms-3">
                                    Elibrary
                                  </div>
                                </a>
                              </li>


                              {this.state.IsUserRepository.includes("Auditor") || this.state.IsUserRepository.includes("ISCG Staff Admin") || this.state.IsUserRepository.includes("ISCG Staff") ?

                                <li className="menu-item ">
                                  <a
                                    // href="https://adibcloud.sharepoint.com/sites/Shariarepository-uat" target="_blank" data-interception="off"
                                    onClick={() => this.openREpository()}
                                    className="menu-link ps-2 text-decoration-none"
                                  >
                                    <div className="p-1  rounded">
                                      <img
                                        src={
                                          this.props.siteUrl +
                                          "/siteAssets/assets/images/images-folder/sharisRepsideNav.png"
                                        }
                                        className="submenu-iconz"
                                      />
                                    </div>
                                    <div className="menu-title bliss2M fontsize-16 ms-4">
                                      Sharia Repository
                                    </div>
                                  </a>
                                </li>
                                : null}
                            </ul>

                          </li>

                          {this.state.ISUSerSouqAlMaal.includes("Maker - SAM") || this.state.ISUSerSouqAlMaal.includes("Checker - SAM") || this.state.ISUSerSouqAlMaal.includes("Master Data Admin - SAM") ?

                            <li className="menu-item">
                              <a
                                // href="javascript:;"
                                onClick={() => this.openAlusool()}
                                className="menu-link p-2 ms-2">
                                <div className="parent-icon shadow p-1 shadow rounded">
                                  <img
                                    src={
                                      this.props.siteUrl +
                                      "/siteAssets/assets/images/images-folder/Soual-Al-Maal.png"
                                    }
                                    className="h-26 w-26"
                                  />
                                </div>
                                <div className="menu-title bliss2M fontsize-18 ms-4">
                                  Souq Al Maal
                                </div>
                              </a>
                            </li>
                            : null}
                          {this.state.IsUserLead.includes("Leads Coordinator") || this.state.IsUserLead.includes("Offer Manager") || this.state.IsUserLead.includes("Departmental SPOC") || this.state.IsUserLead.includes("Agent")
                            || this.state.IsUserLead.includes("Master Data Admin - LeadX") ?
                            <li className="menu-item">
                              <a
                                //  href="https://adibcloud.sharepoint.com/sites/leadmanagement-uat" target="_blank" data-interception="off"
                                onClick={() => this.openLead()}
                                className="menu-link p-2 ms-2">
                                <div className="parent-icon shadow p-1 shadow rounded">
                                  <img
                                    src={
                                      this.props.siteUrl +
                                      "/siteAssets/assets/images/images-folder/LeadxsideNav.png"
                                    }
                                    className="h-30 w-31"
                                  />
                                </div>
                                <div className="menu-title bliss2M fontsize-18">
                                  Lead'x
                                </div>

                              </a>

                            </li>
                            : null}

                        </ul>

                      </div>

                    </div>

                  </div>

                </div>

              </div>

            </aside>

           

            <main className="page-content bg-secondary">
              <div>
                <div className="card background-BlueImg rounded-4">
                  <div className="card-body">
                    <div className="row align-items-center py-1">
                      <div className="px-4 pt-3 pb-0 arrow-on-hover d-md-flex   justify-content-between banner-texxt">
                        <h5 className="mt-0 fontsize-24 text-white ms-2 bliss2M mb-1 text-center text-md-start ">User Access Management</h5>
                        <h5 className="mt-0 fontsize-38 text-white ms-2   mb-1 text-center text-md-end me-2">إدارة وصول المستخدم</h5>
                      </div>
                      <div className="px-4 pt-1 pb-3 arrow-on-hover d-md-flex justify-content-between banner-texxt">
                        <p className="mb-0 fontsize-20  text-white ms-2 text-center text-md-start">Home / User Access Management</p>
                        <p className="mb-0 fontsize-27 text-white ms-2 text-center text-md-end me-2">الرئيسية / إدارة وصول المستخدم</p>
                      </div>  
                    </div> 
                  </div>




                  {/* <div className="card-body">
                    <div className="row align-items-center py-4">
                      <div className=" col-lg-12 px-4 arrow-on-hover d-md-flex justify-content-between banner-texxt">
                        <h5 className="mt-0 ps-4 fontsize-24 text-white ms-0 bliss2M mb-2">User Access Management</h5>
                        <h5 className="mt-0 pe-4 fontsize-24 text-white ms-2 bliss2M text-end mb-2">إدارة وصول المستخدم</h5>
                      </div>
                      <div className=" col-lg-12 p-5 py-4 arrow-on-hover d-flex justify-content-between banner-texxt">
                      <p className="mb-0 ps-4 fontsize-20 text-white ms-1">Home / User Access Management</p>
                        
                        <p className="mb-0 pe-4 fontsize-18 text-white ms-2 text-end">الرئيسية / إدارة وصول المستخدم</p>
                      </div>
                    </div>
                  </div> */}

                </div>

                {
                  (this.state.IsUseradmin.includes('User Admin') === true)
                  && (
                    <div className="row align-items-center">
                      <div className="col-md-6 pt-4">
                        <h5 className="bliss2M font-18 mb-0">All Users</h5>
                      </div>

                      <div className="">

                        <div className="col-md-12 filter-box">
                          <div className="card p-4 mt-2 rounded-4 border-0 filter-inputs">
                            <div className=" row mb-0 p-15" id="myCheck">
                              <div className="col-md-4 col-lg-3 py-2">
                                <Select
                                  value={this.state.selectedId}
                                  onChange={(e) => {
                                    this.handleChangeId(e);
                                  }}
                                  options={this.state.options}
                                  onInputChange={this.handleInputChangeId}
                                  menuIsOpen={this.state.searchId.length > 0}
                                  isMulti={true}
                                  placeholder={"Search EMP ID"}
                                  components={{ DropdownIndicator: () => null, IndicatorSeparator: () => null }}

                                />
                              </div>
                              <div className="col-md-4 col-lg-3 py-2">
                                <Select
                                  value={this.state.selectedName}
                                  onChange={(e) => {
                                    this.handleChangeName(e);
                                  }}
                                  options={this.state.options}
                                  onInputChange={this.handleInputChangeName}
                                  menuIsOpen={this.state.searchName.length > 0}
                                  isMulti={true}
                                  placeholder={"Search Name"}
                                  components={{ DropdownIndicator: () => null, IndicatorSeparator: () => null }}

                                />
                              </div>
                              <div className="col-md-4 col-lg-3 py-2">
                                <Select
                                  value={this.state.selectedEmail}
                                  onChange={(e) => {
                                    this.handleChangeEmail(e);
                                  }}
                                  options={this.state.options}
                                  onInputChange={this.handleInputChangeEmail}
                                  menuIsOpen={this.state.searchEmail.length > 0}
                                  isMulti={true}
                                  placeholder={"Search E-mail ID"}
                                  components={{ DropdownIndicator: () => null, IndicatorSeparator: () => null }}

                                />
                              </div>
                              <div className="col-md-4 col-lg-3 py-2">
                                <Select
                                  value={this.state.SelectedDeptOptions}
                                  // (option: { value: any }) =>
                                  //   option.value === this.state.Depti)}
                                  onChange={(e) => { this.handleChangeDept(e) }}
                                  options={this.state.DeptOptions}
                                  onInputChange={this.handleInputChangeDept}
                                  menuIsOpen={this.state.SearchDept.length > 0}
                                  isMulti={true}
                                  placeholder="Search Department"
                                  components={{ DropdownIndicator: () => null, IndicatorSeparator: () => null }}

                                />
                              </div>
                              <div className="col-md-4 col-lg-3 py-2">
                                <Select
                                  value={this.state.selectedRole}
                                  onChange={(e) => {
                                    this.handleChangeRole(e);
                                  }}
                                  options={this.state.options}
                                  onInputChange={this.handleInputChangeRole}
                                  menuIsOpen={this.state.searchRole.length > 0}
                                  isMulti={true}
                                  placeholder={"Search Role"}
                                  components={{ DropdownIndicator: () => null, IndicatorSeparator: () => null }}

                                />
                              </div>

                              <div className=" col-md-4 col-lg-3 py-2">
                                <Select
                                  value={this.state.iscreatedby}
                                  onChange={(e) => {
                                    this.handleChangecreate(e);
                                  }}
                                  options={this.state.options}
                                  onInputChange={this.handleInputChangecreate}
                                  menuIsOpen={this.state.searchcreate.length > 0}
                                  isMulti={true}
                                  placeholder={"Search Created By"}
                                  components={{ DropdownIndicator: () => null, IndicatorSeparator: () => null }}


                                />
                              </div>
                              <div className="col-md-4 col-lg-3 py-2">
                                <RangePicker
                                  className="form-control bliss2R text-black"
                                  placeholder={['Start Date', 'End Date']}
                                  value={this.state.selectedRange}
                                  onChange={this.handleDateRangeChange}
                                  format={"DD/MM/YYYY"}


                                />
                              </div>

                              {/* <div className="col-md-4 col-lg-3 py-2">
           <Select
             value={this.state.statusOptions.find(
               (option: { value: any }) =>
                 option.value === this.state.statusi
             )}
             onChange={this.handleChangeStatus}
             options={this.state.statusOptions}
             placeholder="Choose Status"
           />
         </div> */}
                            </div>
                          </div>
                        </div>

                      </div>

                    </div>
                  )}

                {
                  (this.state.IsUseradmin && this.state.IsUseradmin.length > 0 && this.state.IsUseradmin.includes('User Admin') === true) ? (
                    <div >

                      <div className="row">
                        <div className="col ">
                          <div className="dataTables_length" id="datatable1_length">
                            <label className="d-flex align-items-center font-17 text-black ">
                              Show
                              <select
                                name="datatable1_length"
                                aria-controls="datatable1"
                                className="form-control w-auto m-3 text-center"
                                value={this.state.itemsPerPage}
                                onChange={(e) =>
                                  this.setState({
                                    itemsPerPage: parseInt(e.target.value),
                                    currentPage: 1,
                                  })
                                }
                              >
                                {this.state.itemsPerPageOptions.map((option: any) => (
                                  <option key={option} value={option}>
                                    {option}
                                  </option>
                                ))}
                              </select>
                              entries
                            </label>
                          </div>
                        </div>
                        <div className="col text-end add-user-col">
                          <button
                            id="Close-Modal"
                            type="button"
                            className="btn btn-primary text-white w-150 my-3"
                            data-bs-toggle="modal"
                            data-bs-target="#Add-User"
                            onClick={() =>
                              this.setState({ onAddhide: true, onUpdatehide: false })
                            }
                          >
                            Add User
                          </button>

                          <div
                            className="modal fade"
                            id="Add-User"
                            aria-labelledby="exampleModalLabel"
                            aria-hidden="true"
                          >
                            <div className="modal-dialog add-user-modal modal-lg">
                              <div className="modal-content">
                                <div className="modal-header bg-primary-header">
                                  {this.state.onAddhide && (
                                    <h5
                                      className="modal-title text-white bliss2M"
                                      id="exampleModalLabel"
                                    >
                                      Add User
                                    </h5>
                                  )}
                                  {this.state.onUpdatehide && (
                                    <h5
                                      className="modal-title text-white bliss2M"
                                      id="exampleModalLabel"
                                    >
                                      Update User
                                    </h5>
                                  )}

                                  <button
                                    type="button"
                                    className="btn-close btn-close-white"
                                    data-bs-dismiss="modal"
                                    aria-label="Close"
                                    onClick={(e) =>
                                      this.setState({
                                        employeeID: "",
                                        empId: "",
                                        userName: "",
                                        emailId: "",
                                        applicationSelectedRoles: [],
                                        applicationSelectedValue: [],
                                        roles: [],
                                        isActive: true,
                                        peopleval: null,
                                        onAddhide: false,
                                        onUpdatehide: false,
                                        valname: false,
                                        valempid: false,
                                        valrole: false,
                                        valmodule: false,
                                      })
                                    }
                                  ></button>
                                </div>
                                <div className="modal-body text-start p-3 peoplepicker_input">
                                  {this.state.onAddhide && (
                                    <div className="mb-2 mt-2">
                                      <label
                                        htmlFor="Add-value-ID"
                                        className="form-label bliss2M font-18 text-dark"
                                      >
                                        Search By Employee Email ID/Employee Name
                                      </label>
                                      <PeoplePicker
                                        context={this.props.context}
                                        personSelectionLimit={1}
                                        groupName={""}
                                        showtooltip={true}
                                        required={true}
                                        placeholder={"Search Employee Email"}
                                        // onChange={(e) =>
                                        //   this.onChangePeoplePickerItems(e)
                                        // } 
                                        onChange={(e) => {
                                          this.onChangePeoplePickerItems(e, this.state.branch);
                                          const email = e[0]?.secondaryText;
                                          let valname = false;

                                          if (!email) {
                                            valname = true;
                                          }

                                          this.setState({ valname });
                                        }}
                                        showHiddenInUI={false}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={1000}
                                        ensureUser={true}
                                        aria-label="Recipient's username"
                                        aria-describedby="basic-addon2"
                                      ></PeoplePicker>
                                      {this.state.valname && (
                                        <p style={{ color: "red" }}>
                                          Please search Employee Email ID
                                        </p>
                                      )}
                                    </div>
                                  )}
                                  {this.state.onAddhide && (
                                    <div className="mb-2">
                                      <label
                                        htmlFor="Add-User-Email"
                                        className="form-label bliss2M font-18 text-dark"
                                      >
                                        Employee ID
                                      </label>
                                      <input
                                        type="number"
                                        id="Add-User-EmpId"
                                        className="form-control w-100 "
                                        aria-describedby="empidHelp"
                                        placeholder="Enter Emp ID"
                                        value={this.state.empId}
                                        // onChange={(e) =>
                                        //   this.setState({ empId: e.target.value }, () => this.Validation())
                                        // }
                                        onChange={(e) => {
                                          const empId = e.target.value;
                                          let valempid = false;

                                          if (!empId) {
                                            valempid = true;
                                          }

                                          this.setState({ empId, valempid });
                                        }}
                                      />
                                      {this.state.valempid && (
                                        <p style={{ color: "red" }}>
                                          Please enter Employee ID
                                        </p>
                                      )}
                                    </div>
                                  )}
                                  <div className="row">
                                    <div className="col-6">
                                      <label
                                        htmlFor="Add-User-Name"
                                        className="form-label bliss2M font-18 text-dark"
                                      >
                                        Name
                                      </label>
                                      <input
                                        type="text"
                                        className="form-control w-100 bg-secondary-5"
                                        id="Add-User-Name"
                                        aria-describedby="emailHelp"
                                        placeholder="Name"
                                        disabled
                                        value={this.state.userName}
                                        onChange={(e) =>
                                          this.setState({ userName: e.target.value })
                                        }
                                      />
                                    </div>

                                    <div className="col-6">
                                      <label
                                        htmlFor="Add-User-Dept"
                                        className="form-label bliss2M font-18 text-dark"
                                      >
                                        Department
                                      </label>
                                      <input
                                        type="text"
                                        className="form-control w-100 bg-secondary-5"
                                        id="Add-User-Dept"
                                        aria-describedby="emailHelp"
                                        placeholder="Department"
                                        disabled
                                        value={this.state.userName === '' ? '' : (this.state.branch)}
                                        onChange={(e) => this.setState({ branch: e.target.value })}
                                      />
                                    </div>
                                  </div>

                                  <div className="mb-2 mt-2">
                                    <label
                                      htmlFor="Add-User-Email"
                                      className="form-label bliss2M font-18 text-dark"
                                    >
                                      Email ID
                                    </label>
                                    <input
                                      type="email"
                                      className="form-control w-100 bg-secondary-5"
                                      id="Add-User-Email"
                                      aria-describedby="emailHelp"
                                      placeholder="Email ID"
                                      disabled
                                      value={this.state.emailId}
                                      onChange={(e) =>
                                        this.setState({ emailId: e.target.value })
                                      }
                                    />
                                    <label
                                      hidden={this.state.isActive}
                                      style={{ color: "red" }}
                                    >
                                      Email ID Already Exists
                                    </label>
                                  </div>

                                  <div className="row add-user-focusremove">
                                    <div className="col-6">
                                      <label
                                        htmlFor="Add-value-Description"
                                        className="form-label bliss2M font-18 text-dark"
                                      >
                                        Module
                                      </label>
                                      <Select
                                        value={this.state.applicationSelectedValue}
                                        options={this.state.values}
                                        onChange={(selectedOption) => {
                                          this.handleapplication(selectedOption);
                                          const valmodule = !selectedOption;
                                          this.setState({ valmodule });

                                        }}
                                      />
                                      {this.state.valmodule && (
                                        <p style={{ color: "red" }}>
                                          Please select Module
                                        </p>
                                      )}
                                    </div>
                                    <div className="col-6 usc-adduser-btn">
                                      <label
                                        htmlFor="Add-User-AssignRole"
                                        className="form-label bliss2M font-18 text-dark"
                                      >
                                        Assign Role
                                      </label>
                                      <Select
                                        value={this.state.applicationSelectedRoles}
                                        // options={this.state.newOption}
                                        options={this.state.roles}
                                        onChange={(selectedOptions) => {
                                          this.handleroles(selectedOptions);
                                         // console.log(selectedOptions);
                                          const valrole = selectedOptions.length === 0 && !this.state.onUpdatehide;
                                          this.setState({ valrole });
                                        }}
                                        isMulti={true}
                                        isDisabled={!this}
                                      // onInputChange={(e)=>{
                                      //   console.log(e);
                                      // }}
                                      />
                                      {this.state.valrole && (
                                        <p style={{ color: "red" }}>
                                          Please select Assign Role
                                        </p>
                                      )}
                                    </div>
                                  </div>
                                </div>
                                <div className="modal-footer">
                                  <button
                                    type="button"
                                    className="btn btn-primary text-white w-150"
                                    data-bs-dismiss="modal"
                                    onClick={(e) =>
                                      this.setState({
                                        empId: "",
                                        userName: "",
                                        emailId: "",
                                        branch: "",
                                        applicationSelectedRoles: [],
                                        applicationSelectedValue: [],
                                        roles: [],
                                        isActive: true,
                                        peopleval: null,
                                        onAddhide: false,
                                        onUpdatehide: false,
                                        valname: false,
                                        valempid: false,
                                        valmodule: false,
                                        valrole: false,
                                      }
                                      )
                                    }
                                  >
                                    Cancel
                                  </button>
                                  {this.state.onAddhide && (
                                    <button
                                      type="button"
                                      className="btn btn-primary text-white w-150"
                                      // data-bs-dismiss="modal"
                                      onClick={(e) => this.Add(e)}
                                    >
                                      Add
                                    </button>
                                  )}
                                  {this.state.onUpdatehide && (
                                    <button
                                      type="button"
                                      className="btn btn-primary text-white w-150"
                                      // data-bs-dismiss="modal"
                                      onClick={(e) => this.update(e)}
                                    >
                                      Update
                                    </button>
                                  )}
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div
                        className="modal fade"
                        id="staticBackdrop"
                        data-bs-backdrop="static"
                        data-bs-keyboard="false"
                        aria-labelledby="staticBackdropLabel"
                        aria-hidden="true"
                      >
                        <div className="modal-dialog modal-lg">
                          <div className="modal-content">
                            <div className="modal-header border-0">
                              <button
                                type="button"
                                className="btn-close"
                                data-bs-dismiss="modal"
                                aria-label="Close"
                              ></button>
                            </div>
                            <div className="modal-body text-center">
                              <img
                                src={
                                  this.props.context.pageContext.web.absoluteUrl +
                                  "/siteAssets/assets/images/images-folder/deleteicon-MasterData.png"
                                }
                                className="p-2 img-fluid w-64 h-64"
                              />
                              {/* <h2 className="p-2 font-30 mt-2">Are You Sure ?</h2> */}
                              <p className="text-secondary p-2 font-22">
                                You want to Delete this User record ?
                              </p>
                            </div>
                            <div className="modal-footer border-0">
                              <button
                                type="button"
                                className="btn btn-primary text-white w-150"
                                data-bs-dismiss="modal"
                              >
                                Cancle
                              </button>
                              <button
                                type="button"
                                className="btn btn-primary text-white w-150 my-3"
                                data-bs-dismiss="modal"
                                onClick={(e: any) => this.handleCheck(this.state.DeleteId)}
                              // , this.state.DeleteStatus
                              >
                                Yes, Delete
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="Container-fluid read-item-full-table">
                        <div className="rounded-4 table-responsive shadow read-item-table">
                          <table
                            className="table table-striped rounded-4 bg-white mb-0"
                            id="table-useraccess"
                          >
                            <thead className="table-primary text-left font-17 ">
                              <tr>
                                <th className="fw-normal px-4 col">Emp ID</th>
                                <th className="fw-normal col">Name</th>
                                <th className="fw-normal col ">E-Mail ID</th>
                                {/* <th className="fw-normal ">Module</th> */}
                                <th className="fw-normal col ">Role Name</th>
                                <th className="fw-normal col">Department</th>
                                <th className="fw-normal col" hidden>
                                  {" "}
                                  Created By{" "}
                                </th>
                                <th className="fw-normal text-right">Actions</th>
                              </tr>
                            </thead>
                            <tbody className=" font-16 ">
                              {currentData &&
                                currentData.length > 0 ?
                                currentData.map((obj: any) => {
                                  return (
                                    <tr>
                                      <td className="text-left px-4 py-4">
                                        {obj.employeeId}
                                      </td>
                                      <td className="text-left py-4">
                                        {obj.userName}
                                      </td>
                                      <td className="text-left py-4">
                                        {obj.emailId}
                                      </td>
                                      {/* <td className="text-left py-4">

                                        {obj.roles.map((role: any, i: any) => {

                                          var rol = "";
                                          if (i >= 1) {
                                            rol =
                                              rol +
                                              rol.concat(",", String(role.applicationName));
                                          } else {
                                            rol = rol.concat(String(role.applicationName));
                                          }
                                          return rol;
                                        })}
                                      </td> */}

                                      {/* <td className="text-left py-4">
                                        {obj.roles.map((role: any, i: any) => {
                                          var rol = "";
                                          if (i >= 1) {
                                            rol =
                                              rol +
                                              rol.concat(",", String(role.name));
                                          } else {
                                            rol = rol.concat(String(role.name));
                                          }
                                          return rol;
                                        })}
                                      </td> */}
                                      {/* <td className="text-left py-4">
                                        {obj.roles.map((role: any, i: any) => {
                                          var roleName = role.name; 
                                          if (roleName === "User Admin") {
                                            roleName = "UAM Admin";
                                          }
                                          if (roleName === "Auditor") {
                                            roleName = "Sharia Repository - Auditor";
                                          }
                                          if (roleName === "ISCG Staff") {
                                            roleName = "Sharia Repository - ISCG Staff";
                                          }
                                           if (roleName === "ISCG Staff Admin") {
                                            roleName = "Sharia Repository - Master Data Manager";
                                          }
                                          if (roleName === "ELibrary Admin") {
                                            roleName = "Alif E-Library – Master Data Manager";
                                          }
                                          if (roleName === "ELibrary-ISCG Staff") {
                                            roleName = "Alif E-Library - ISCG Staff";
                                          }
                                          if (roleName === "Helpdesk Manager-Sharia Research") {
                                            roleName = "Sharia Research - Helpdesk Manager";
                                          }
                                          if (roleName === "ISCG Staff-ShariaResearch") {
                                            roleName = "Sharia Research - ISCG Staff";
                                          }
                                          if (roleName === "Master Data Admin-ShariaResearch") {
                                            roleName = "Sharia Research - Master Data Manager";
                                          }
                                          var rol = "";
                                          if (i >= 1) {
                                            rol = rol + rol.concat(",", String(roleName));
                                          } else {
                                            rol = rol.concat(String(roleName));
                                          }
                                          return rol;
                                        })}
                                      </td> */}
                                      {/* <td className="text-left py-4">
                                        {obj.roles.map((role: any, i: any) => {
                                          const roleName = this.mapRoleName(role.name, );

                                          if (i >= 1) {
                                            return roleName;
                                          } else {
                                            return roleName;
                                          }
                                        })}
                                      </td> */}
                                      <td className="text-left py-4">
                                        {obj.roles.map((role: any, i: any) => {
                                          const roleName = this.mapRoleName(role.name);
                                          if (i >= 1) {
                                            return `,  ${roleName}`;

                                          } else { return roleName; }
                                        })} </td>
                                      <td className="text-left py-4">
                                        {obj.department}
                                      </td>
                                      <td className="text-left py-4" hidden>
                                        {obj.createdby}
                                      </td>
                                      <td>
                                        <div className="d-flex text-right pt-2 px-4 icon-boxxx py-4">
                                          {obj.status && (
                                            <img
                                              src={
                                                this.props.context.pageContext.web
                                                  .absoluteUrl +
                                                "/siteAssets/assets/images/images-folder/MasterPage-EditIcon.svg"
                                              }
                                              className="edit-iconn cursor-pointer"
                                              width="20"
                                              data-bs-toggle="modal"
                                              data-bs-target="#Add-User"
                                              onClick={() => this.getdatabyid(obj.id)}
                                            />
                                          )}
                                          &nbsp;&nbsp;&nbsp;&nbsp;
                                          <a
                                            href=""
                                            className="modal-dialog-centered"
                                            data-bs-toggle="modal"
                                            data-bs-target="#staticBackdrop"
                                          >
                                            {obj.status && (
                                              <img
                                                src={
                                                  this.props.context.pageContext.web
                                                    .absoluteUrl +
                                                  "/siteAssets/assets/images/images-folder/MasterPage-deleteIcon.svg"
                                                }
                                                className="delete-iconn "
                                                width="17"
                                                onClick={(e: any) =>
                                                  this.setState({
                                                    DeleteId: obj.id,
                                                    DeleteStatus: obj.status
                                                  })
                                                }

                                              />
                                            )}
                                          </a>
                                        </div>
                                      </td>
                                    </tr>
                                  );
                                }) :
                                // <label className="text-center no-userfound ">No Users Found</label>

                                <tr>
                                  <td className="text-center py-4" colSpan={6}>
                                    <label className="no-userfound">No Users Found</label>
                                  </td>
                                </tr>
                              }

                            </tbody>
                          </table>
                        </div>
                      </div>
                      <div className=" help-desk-Homepage mt-4">
                        <div className="row pt-3">
                          {/* <div className="col-xs-12 col-sm-12 col-md-6 col-lg-6 pt-2 show-pagex-entries">
                            <p className="text-black font-17 ">Showing 1 to 05 entries of 200</p>
                          </div> */}
                          <div className="col-xs-12 col-sm-12 col-md-6 col-lg-6 pt-2 show-pagex-entries">
                            <p className="text-black font-17">
                              Showing {this.getStartIndex()} to {this.getEndIndex()} entries of {this.state.datarecordsDup.length}
                            </p>
                          </div>
                          <div className="col-xs-12 col-sm-12 col-md-6 col-lg-6 ">
                            <nav aria-label="...">
                              <ul className="pagination pagination-md custom-pagination justify-content-end mb-3">
                                {/* <li className="mx-2 text-center"><a className="page-link text-dark rounded">
                     <i
                     className="bi bi-chevron-left" onClick={() => { if (this.state.currentPage - 3 > 0) { this.setState({ pagenumber: this.state.pagenumber - 3 }, () => { this.setState({ currentPage: this.state.pagenumber }) }) } }} ></i></a></li>
               
                   <li className="mx-2 text-center"><a className="page-link text-dark rounded"><i
                     className="bi bi-chevron-right" onClick={() => { if (this.state.currentPage < this.totalPages) { this.setState({ pagenumber: this.state.pagenumber + 3 }, () => { this.setState({ currentPage: this.state.pagenumber }) }) } }}>
                       </i></a></li> */}
                                {this.renderpagination()}
                              </ul>
                            </nav>
                          </div>
                        </div>
                      </div>
                    </div>
                  ) :
                    (this.state.IsUseradmin && this.state.IsUseradmin.length > 0 && this.state.IsUseradmin.includes('User Admin') !== true) ?
                      <div className="text-center font-35 pt-5">You don't have access to view this Page.</div> : null
                }


              </div>
            </main>
            <footer className=" bg-secondary-2 footer-custom h-55 px-3">
              <p className="text-center text-white pt-3 font-16 ">
                Copyrights Reserved Abu Dhabi Islamic Public Joint Stock Company
                2022
              </p> 
            </footer> 
            <div className="overlay nav-toggle-icon"></div>
          </div>
        </body>
      </section>
    );
  }


  public mapRoleName(roleName: any) {
    switch (roleName) {
      case "User Admin":
        return "UAM MasterData - Manager";
      case "Auditor":
        return "Sharia Repository - Auditor";
      case "ISCG Staff":
        return "Sharia Repository - ISCG Staff";
      case "ISCG Staff Admin":
        return "Sharia Repository - Master Data Manager";
      case "ELibrary Admin":
        return "Alif E-Library – Master Data Manager";
      case "ELibrary-ISCG Staff":
        return "Alif E-Library - ISCG Staff";

      case "Helpdesk Manager-Sharia Research":
        return "Sharia Research - Helpdesk Manager";

      case "ISCG Staff-ShariaResearch":
        return "Sharia Research - ISCG Staff";
      case "Master Data Admin-ShariaResearch":
        return "Sharia Research - Master Data Manager";
      case "Helpdesk Manager":
        return "Sharia Helpdesk - Helpdesk Manager";
      case "ISCG Staff - Helpdesk":
        return "Sharia Helpdesk - ISCG Staff";
      case "Master Data Admin - Helpdesk":
        return "Sharia Helpdesk - Master Data Manager";
      default:
        return roleName;
    }
  }


  ////Token Generation

  public getAccessToken = async (email: string) => {
    try {
      debugger;
      sessionStorage.clear();
      localStorage.clear();
      // console.log(accounts);
      // this.GetIdTokenClaims();
      // let accessToken: string = null;
      this.setCurrentAccount(email);//Sets the current account.
      //console.log(currentAccount);
      // this.setState({ isLoading: true });
      const loginRedirect: msal.RedirectRequest = this.state.tokenrequest;
      loginRedirect.loginHint = email;
      await this.state.msalInstance.acquireTokenPopup(loginRedirect)
        .then((tokenResponse: any) => {

          // console.log(tokenResponse);
          //    this.setState({ isLoading: false });
          localStorage.setItem("AccessToken", tokenResponse.accessToken);
          //console.log(tokenResponse.accessToken);
          this.setState({ accessToken: tokenResponse.accessToken }, () => {
            // console.log(this.state.accessToken)
            this.allUser();
            this.Application();
            this.Search();
            this.Roles();
            this.findUser();
            this.GetApplicationModules();
          });



          // Constants.accessToken = tokenResponse.accessToken;
          // Constants.homeAccountId = tokenResponse.account.homeAccountId;
          // console.log("Token: ", Constants.accessToken);
          // return tokenResponse.accessToken;
        })
        // .then((tokenResponse)=>{
        //     console.log(tokenResponse);
        //     localStorage.setItem("AccessToken",tokenResponse)

        // })
        .catch((error: any) => {
          console.error(error);
          this.interactionRequired(email);


          // I haven't implemented redirect but it is fairly easy
          console.error("Maybe it is a popup blocked error. Implement Redirect");
          //return null;

        });

    } catch (error) {
      console.error(error);
    }
  };

  protected setCurrentAccount(email: string) {
    try {
      const currentAccounts: msal.AccountInfo[] = this.state.msalInstance.getAllAccounts();
      //returns all the accounts currently in the cache. An application must choose an account to acquire tokens silently.
      if (currentAccounts === null || currentAccounts.length === 0) {
        currentAccount = this.state.msalInstance.getAccountByUsername(
          email
        );
        //receives a username string and returns the matching account from the cache.
      } else if (currentAccounts.length > 1) {
        console.warn("Multiple accounts detected.");
        currentAccount = this.state.msalInstance.getAccountByUsername(
          email
        );
      } else if (currentAccounts.length === 1) {
        currentAccount = currentAccounts[0];
      }
      this.state.tokenrequest.account = currentAccount;
    } catch (error) {
      console.error(error);
    }
  };

  protected interactionRequired = (email: string): Promise<string> => {
    try {
      // this.setState({ isLoading: true });
      //console.log("Inside Interaction");
      const loginPopupRequest: msal.PopupRequest = this.state.tokenrequest;
      loginPopupRequest.loginHint = email;
      return this.state.msalInstance
        //The resolution is to call an interactive method such as acquireTokenPopup
        .acquireTokenPopup(loginPopupRequest)
        .then((tokenResponse: any) => {
          //   this.setState({ isLoading: false });
          localStorage.setItem("AccessToken", tokenResponse.accessToken);
          // Constants.accessToken = tokenResponse.accessToken;
          // Constants.homeAccountId = tokenResponse.account.homeAccountId;
          return tokenResponse.accessToken;

        })
        .catch((error: any) => {
          console.error(error);
          // I haven't implemented redirect but it is fairly easy
          console.error("Maybe it is a popup blocked error. Implement Redirect");
          // return null;
        });
    } catch (error) {
      console.error(error);
    }
  };

  get totalPages() {
    return Math.ceil(this.state.datarecords.length / this.state.itemsPerPage);
  }

  renderpagination = () => {
    const { currentPage, itemsPerPage, datarecordsDup } = this.state;
    const totalPages = Math.ceil(datarecordsDup.length / itemsPerPage);
    const pages = [];
    let startPage = Math.max(1, currentPage - 1);
    let endPage = Math.min(startPage + 2, totalPages);
    if (endPage - startPage < 2) {
      startPage = Math.max(1, endPage - 2);
    }
    for (let i = startPage; i <= endPage; i++) {
      if (i === currentPage) {
        pages.push(
          <li key={i} className="disabled text-center cursor-pointer">
            <a className="page-link text-white btn-primary rounded px-3 active">{i}</a>
          </li>
        );
      } else {
        pages.push(
          <li key={i} className=" text-center">
            <a className="page-link text-dark rounded" onClick={() => this.setState({ currentPage: i })}>{i}</a>
          </li>
        );
      }
    }
    return (
      <ul className="pagination">
        {currentPage !== 1 && (
          <>
            <li className="text-center">
              <a className="page-link text-dark rounded" onClick={() => this.setState({ currentPage: 1 })}>
                <i className="bi bi-chevron-double-left"></i>
              </a>
            </li>
            <li className=" text-center">
              <a className="page-link text-dark rounded" onClick={() => this.setState({ currentPage: currentPage - 1 })}>
                <i className="bi bi-chevron-left"></i>
              </a>
            </li>
          </>
        )}
        {pages}
        {currentPage !== totalPages && (
          <>
            <li className=" text-center">
              <a className="page-link text-dark rounded" onClick={() => this.setState({ currentPage: currentPage + 1 })}>
                <i className="bi bi-chevron-right"></i>
              </a>
            </li>
            <li className="text-center">
              <a className="page-link text-dark rounded" onClick={() => this.setState({ currentPage: totalPages })}>
                <i className="bi bi-chevron-double-right"></i>
              </a>
            </li>
          </>
        )}
      </ul>
    );
  }
  getStartIndex = () => {
    const { currentPage, itemsPerPage } = this.state;
    return (currentPage - 1) * itemsPerPage + 1;
  };

  getEndIndex = () => {
    const { currentPage, itemsPerPage, datarecordsDup } = this.state;
    const endIndex = Math.min(currentPage * itemsPerPage, datarecordsDup.length);
    return endIndex;
  };






  public handleCheck = (id: any) => {
    try {
      // if (status == true) {
      axios(
        `${Constants.var}User/Delete?id=${id}&deletedby=${this.props.context.pageContext.user.email}`,
        {
          method: "put",
          headers: { Authorization: `Bearer ${this.state.accessToken}` }
        }
      )
        .then((response) => { })
        .then(() => {
          // this.filterGrid();
          this.Search();
          this.allUser();

        })

      // .then(() => this.allUser())
      // .catch((error) => { });
      // } else {
      //   this.Search();
      //   this.allUser();
      // }
      // if (status == false) {
      //   axios(
      //     `${Constants.var}/api/User/Delete?id=${id}&deletedby=${this.props.context.pageContext.user.email}`,
      //     {
      //       method: "put",
      //     }
      //   )
      //     .then((response) => { })
      //     .then(() => {
      //       this.Search();
      //       this.allUser();
      //     })
      //     .then(() => this.allUser())
      //     .catch((error) => { });
      //   this.Search();
      // } else {
      // }
    } catch (exception) { }
  };
  public handleapplication = (applicationSelectedValue: any) => {
    this.setState({ applicationSelectedValue, roles: [] }, () => {
      this.Roles();
      if (this.state.applicationSelectedValue.length == 0) {
        this.setState({
          valrole: true,
        })
      }
    });
  };
  public handleroles = (applicationSelectedRoles: any) => {
    this.setState({ applicationSelectedRoles })
    //var a =[];
    const CheckRemover = this.state.applicationSelectedRoles.filter((prevState: any) => !applicationSelectedRoles.some((obj: any) => obj.value === prevState.value))
    CheckRemover.forEach((checkroles: any) => {
      this.setState((prevState) => ({
        RemoveID: [...prevState.RemoveID, checkroles.value]
      }))

    })

    // if (Remover) {
    //   this.setState((prevState) => ({
    //     CheckRemover: [...prevState.CheckRemover, Remover]
    //   }))
    // }
    //  console.log(this.state.CheckRemover,'remover')
    //this.setState({CheckRemover})
  };


  ////siva method///
  GetApplicationModules() {
    let requestURL = `${Constants.var}User/GetApplication`
    // alert(requestURL)
    try {

      this.GetmethodAJAX(requestURL)
        .then((response: any) => {
          if (response && response.length) {
            // let listdata = [];
            for (let i = 0; i <= response.length; i++) {
              let ModuleName: any;
              if (response.length > 0) {
                ModuleName = response[i].applicationName
              }
              if (ModuleName === "Helpdesk") {
                this.setState({ HelpdeskUrl: response[i].aplicationUrl })
                //console.log(this.state.HelpdeskUrl, 'HelpdeskURL');
              }
              else if (ModuleName === "User Access Management") {
                this.setState({ UAMUrl: response[i].aplicationUrl })
                // console.log(this.state.UAMUrl, "User Acess M")
              }
              else if (ModuleName === "Sharia Repository") {
                this.setState({ ShariaRepoUrl: response[i].aplicationUrl })
                // console.log(this.state.ShariaRepoUrl, "Sharia repo")
              }
              else if (ModuleName === "ELibrary") {
                this.setState({ eLibraryUrl: response[i].aplicationUrl })
                // console.log(this.state.eLibraryUrl, "ELibrary")
              }
              else if (ModuleName === "Sharia Research") {
                this.setState({ ShariaResearchUrl: response[i].aplicationUrl })
                // console.log(this.state.ShariaResearchUrl, "shariaresearch")
              }
              else if (ModuleName === "LeadX") {
                this.setState({ LeadXUrl: response[i].aplicationUrl })
                // console.log(this.state.LeadXUrl, "leadx")
              }
              else if (ModuleName === "Tawazun") {
                this.setState({ TawazunUrl: response[i].aplicationUrl })
                // console.log(this.state.TawazunUrl, "Tawazun")
              }
              else if (ModuleName === "Vertex") {
                this.setState({ VertexUrl: response[i].aplicationUrl })
                // console.log(this.state.VertexUrl, "Vertex")
              }
              else if (ModuleName === "Apex") {
                this.setState({ ApexUrl: response[i].aplicationUrl })
                // console.log(this.state.ApexUrl, "Apex")
              }
              else if (ModuleName === "Souq Al Maal") {
                this.setState({ SouqAlMaalUrl: response[i].aplicationUrl })
                // console.log(this.state.SouqAlMaalUrl, "Souq Al Maal")
              }
            }
          }
        })
        .catch((err: any) => {
        });
    } catch (error) {
    }
  }

  public openHelpdesk = () => {
    try {
      window.open(
        `${this.state.HelpdeskUrl}`,
        "_blank"
      );
    } catch (error) { }

  };

  public openShariaReasearch = () => {
    try {
      window.open(
        `${this.state.ShariaResearchUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openTawazun = () => {
    try {
      window.open(
        `${this.state.TawazunUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openapex = () => {
    try {
      window.open(
        `${this.state.ApexUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openvertex = () => {
    try {
      window.open(
        `${this.state.VertexUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openElibrary = () => {
    try {
      window.open(
        `${this.state.eLibraryUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openREpository = () => {
    try {
      window.open(
        `${this.state.ShariaRepoUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openAlusool = () => {
    try {
      window.open(
        `${this.state.SouqAlMaalUrl}`,
        "_blank"
      );
    } catch (error) { }

  };
  public openLead = () => {
    try {
      window.open(
        `${this.state.LeadXUrl}`,
        "_blank"
      );
    } catch (error) { }

  };


  public GetmethodAJAX(URL: any): Promise<any> {
    return new Promise<any>((success, failure) => {
      try {
        $.ajax({
          "url": URL,
          "method": 'GET',
          "headers": {
            "Content-Type": "text/plain",
            Authorization: `Bearer ${this.state.accessToken}`
          },
          success: function (message) {
            success(message);
          },
          error: function (message) {
            failure(message);
          }
        });
      } catch (error) {
        failure(error);
      }
    });
  }


  openpowerBI=()=>
  {
    try {
     
      let a = document.createElement("a");
      a.href =Constants.PowerBiReport

     // console.log(Constants.PowerBiReport,"power bi url")
     
      document.body.appendChild(a);

      a.click();

      document.body.removeChild(a);

      document
        .querySelector("[data-automation-id=contentScrollRegion]")
        .scroll({ top: 0, behavior: "smooth" });
    } catch (ex) { }
  }


public modifydata = (list:any,toupdate:any)=>
{
  var setDropDownValue: any = [];
  if(list.length>0)
  {
   // console.log(list);
    debugger;
  for(let i=0;i<list.length;i++)
  {
    let value=list[i].label;
    let itemvalue=list[i].value;
    switch(value)
    {
      case "User Admin":setDropDownValue.push({label:'UAM - Master Data Manager',value:itemvalue});break;
      case "Auditor":setDropDownValue.push({label:'Sharia Repository - Auditor',value:itemvalue});break;
      case "ISCG Staff":setDropDownValue.push({label:'Sharia Repository - ISCG Staff',value:itemvalue});break;
      case "ISCG Staff Admin":setDropDownValue.push({label:'Sharia Repository - Master Data Manager',value:itemvalue});break;
      case "ELibrary Admin":setDropDownValue.push({label:'Alif E-Library – Master Data Manager',value:itemvalue});break;
      case "ELibrary-ISCG Staff":setDropDownValue.push({label:'Alif E-Library - ISCG Staff',value:itemvalue});break;
      case "Helpdesk Manager-Sharia Research":setDropDownValue.push({label:'Sharia Research - Helpdesk Manager',value:itemvalue});break;
      case "ISCG Staff-ShariaResearch":setDropDownValue.push({label:'Sharia Research - ISCG Staff',value:itemvalue});break;
      case "Master Data Admin-ShariaResearch":setDropDownValue.push({label:'Sharia Research - Master Data Manager',value:itemvalue});break;
      case "Helpdesk Manager":setDropDownValue.push({label:'Sharia Helpdesk - Helpdesk Manager',value:itemvalue});break;
      case "ISCG Staff - Helpdesk":setDropDownValue.push({label:'Sharia Helpdesk - ISCG Staff',value:itemvalue});break;
      case "Master Data Admin - Helpdesk":setDropDownValue.push({label:'Sharia Helpdesk - Master Data Manager',value:itemvalue});break;
      case "User Management Admin":setDropDownValue.push({label:'UAM - Master Data Manager',value:itemvalue});break;
      case "SouqAlMaal Master Data Admin":setDropDownValue.push({label:'SouqAlMaal - Master Data Manager',value:itemvalue});break;
      case "Master Data Admin - Apex":setDropDownValue.push({label:'Apex - Master Data Manager',value:itemvalue});break;
      case "Master Data Admin - Tawazun":setDropDownValue.push({label:'Tawazun - Master Data Manager',value:itemvalue});break;
      default:setDropDownValue.push({label:value,value:itemvalue});
     
    }
  }
  if(toupdate==="fromgetbyid"){
  this.setState({applicationSelectedRoles:setDropDownValue})
  }
  else if(toupdate==="setapplicationroles")
  {
    this.setState({roles:setDropDownValue})
  }
  }
}

  public ShowDatainUAM = (e: any) => {
    var setDropDownValue: any = [];
    if (this.state.roles.length > 0 && this.state.roles[0].label != null) {
      const getallroles = this.state.roles.map((obj: any) => obj.label)
      for (let i = 0; i < this.state.roles.length; i++) {
        setDropDownValue.push({
          label: getallroles.includes("User Admin") ? 'UAM MasterData - Manager' :
            getallroles.includes("Auditor") ? 'Sharia Repository - Auditor' :
              getallroles.includes("ISCG Staff") ? 'Sharia Repository - ISCG Staff' :
                getallroles.includes("ISCG Staff Admin") ? 'Sharia Repository - Master Data Manager' :
                  getallroles.includes("ELibrary Admin") ? 'Alif E-Library – Master Data Manager' :
                    getallroles.includes("ELibrary-ISCG Staff") ? 'Alif E-Library - ISCG Staff' :
                      getallroles.includes("Helpdesk Manager-Sharia Research") ? 'Sharia Research - Helpdesk Manager' :
                        getallroles.includes("ISCG Staff-ShariaResearch") ? 'Sharia Research - ISCG Staff' :
                          getallroles.includes("Master Data Admin-ShariaResearch") ? 'Sharia Research - Master Data Manager' :
                            getallroles.includes("Helpdesk Manager") ? 'Sharia Helpdesk - Helpdesk Manager' :
                              getallroles.includes("ISCG Staff - Helpdesk") ? 'Sharia Helpdesk - ISCG Staff' :
                                getallroles.includes("Master Data Admin - Helpdesk") ? 'Sharia Helpdesk - Master Data Manager'
                                  : null,


          value: getallroles.includes("User Admin") ? this.state.roles[i].value :
            getallroles.includes("Auditor") ? this.state.roles[i].value :
              getallroles.includes("ISCG Staff") ? this.state.roles[i].value :
                getallroles.includes("ISCG Staff Admin") ? this.state.roles[i].value :
                  getallroles.includes("ELibrary Admin") ? this.state.roles[i].value :
                    getallroles.includes("ELibrary-ISCG Staff") ? this.state.roles[i].value :
                      getallroles.includes("Helpdesk Manager-Sharia Research") ? this.state.roles[i].value :
                        getallroles.includes("ISCG Staff-ShariaResearch") ? this.state.roles[i].value :
                          getallroles.includes("Master Data Admin-ShariaResearch") ? this.state.roles[i].value :
                            getallroles.includes("Helpdesk Manager") ? this.state.roles[i].value :
                              getallroles.includes("ISCG Staff - Helpdesk") ? this.state.roles[i].value :
                                getallroles.includes("Master Data Admin - Helpdesk") ? this.state.roles[i].value :
                                  null
        })
      }
      this.setState({
        roles: setDropDownValue

      })
    }
    else if (e && e.length > 0 && e[0].roles[0].name != null) {
      const getallroles = e[0].roles.map((obj: any) => obj.name)
      for (let i = 0; i < e.length; i++) {
        setDropDownValue.push({
          label: getallroles.includes("User Admin") ? 'UAM MasterData - Manager' :
            getallroles.includes("Auditor") ? 'Sharia Repository - Auditor' :
              getallroles.includes("ISCG Staff") ? 'Sharia Repository - ISCG Staff' :
                getallroles.includes("ISCG Staff Admin") ? 'Sharia Repository - Master Data Manager' :
                  getallroles.includes("ELibrary Admin") ? 'Alif E-Library – Master Data Manager' :
                    getallroles.includes("ELibrary-ISCG Staff") ? 'Alif E-Library - ISCG Staff' :
                      getallroles.includes("Helpdesk Manager-Sharia Research") ? 'Sharia Research - Helpdesk Manager' :
                        getallroles.includes("ISCG Staff-ShariaResearch") ? 'Sharia Research - ISCG Staff' :
                          getallroles.includes("Master Data Admin-ShariaResearch") ? 'Sharia Research - Master Data Manager' :
                            getallroles.includes("Helpdesk Manager") ? 'Sharia Helpdesk - Helpdesk Manager' :
                              getallroles.includes("ISCG Staff - Helpdesk") ? 'Sharia Helpdesk - ISCG Staff' :
                                getallroles.includes("Master Data Admin - Helpdesk") ? 'Sharia Helpdesk - Master Data Manager' :

                                  null,


          value: getallroles.includes("User Admin") ? e[0].roles[i].roleId :
            getallroles.includes("Auditor") ? this.state.roles[i].roleId :
              getallroles.includes("ISCG Staff") ? this.state.roles[i].roleId :
                getallroles.includes("ISCG Staff Admin") ? this.state.roles[i].roleId :
                  getallroles.includes("ELibrary Admin") ? this.state.roles[i].roleId :
                    getallroles.includes("ELibrary-ISCG Staff") ? this.state.roles[i].roleId :
                      getallroles.includes("Helpdesk Manager-Sharia Research") ? this.state.roles[i].roleId :
                        getallroles.includes("ISCG Staff-ShariaResearch") ? this.state.roles[i].roleId :
                          getallroles.includes("Master Data Admin-ShariaResearch") ? this.state.roles[i].roleId :
                            getallroles.includes("Helpdesk Manager") ? this.state.roles[i].roleId :
                              getallroles.includes("ISCG Staff - Helpdesk") ? this.state.roles[i].roleId :
                                getallroles.includes("Master Data Admin - Helpdesk") ? this.state.roles[i].roleId :
                                  null
        })
      }
      this.setState({
        applicationSelectedRoles: setDropDownValue
      })
    }


   // console.log(setDropDownValue)

  }
}
