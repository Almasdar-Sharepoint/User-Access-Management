export interface IUserAccessManagementState {
    employeeID: string;
    userName: string;
    branch: string;
    emailId: string;
    roles: any;
    datarecords: any;
    values: any;
    isActive: boolean;
    applicationSelectedValue: any;
    applicationSelectedRoles: any;
    CheckRemover:any
    RemoveID:any
    RemovedRoleIds: any
    Userid: any;
    empId: any;
    peopleval: any;
    defaultroles: any;
    pages: any;
    currentPage: any;
    itemsPerPage: any;
    currentData: any;
    itemsPerPageOptions: any[];
    onAddhide: boolean;
    onUpdatehide: boolean;
    valname: boolean;
    valempid: boolean;
    valmodule: boolean;
    valrole: boolean;
    selectedId: any;
    searchId: any;
    options: any;
    selectedName: any;
    searchName: any;
    selectedEmail: any;
    searchEmail: any;
    selectedRole: any,
    searchRole: any;
    Depti: any;
    DeptOptions: any;
    statusi: any;

    // statusOptions: any; 
    iscreatedby: any;
    searchcreate: any;
    datarecordsDup: any;    //added by siva krishna
    SelectedDeptOptions: any;
    SearchDept: any;
    DeleteId: any
    DeleteStatus: any
    validatemodulecheck: boolean,
    pagenumber: any,
    selectedRange: any,
    isDataFound: any,
    IsUseradmin: any,
    IsUserRepository: any
    IsUserLead: any
    IsUserTawazun: any,
    IsUserApex: any
    ISUSerSouqAlMaal: any
    IsUserShariaResearch: any
    IsVertex:any
    role: any,
    EmPIDCheck: any,
    CheckEmployeeID: boolean

    /////siva buri///
    HelpdeskUrl: any
    UAMUrl: any
    ShariaRepoUrl: any
    eLibraryUrl: any
    ShariaResearchUrl: any
    LeadXUrl: any
    TawazunUrl: any
    VertexUrl: any
    ApexUrl: any
    SouqAlMaalUrl: any
    accessToken:any

    msalInstance?: any;
    currentAccount?: any;
    tokenrequest?: any;
    isLoading:any;
    isHide:boolean;
    newOption:any

}


export interface IUserAccessManagementProps {
    // description: string;
    // isDarkTheme: boolean;
    // environmentMessage: string;
    // hasTeamsContext: boolean;
    // userDisplayName: string;
    context: any;
    // branch: any;
    // Constants: any;
    // options: any;
    // pages: any;
    // currentPage: any;
    // itemsPerPage: any;
    // currentData: any;
    // itemsPerPageOptions: any[];

    siteUrl: any,
    email:any,
    AccessToken:any,
    
    
}

