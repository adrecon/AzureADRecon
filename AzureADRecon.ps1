<#

.SYNOPSIS

    AzureADRecon is a tool which gathers information about the Azure Active Directory and generates a report which can provide a holistic picture of the current state of the target environment.

.DESCRIPTION

    AzureADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of an Azure AD environment with a valid credential. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target environment.
    The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
    The tool requires AzureAD PowerShell Module to be installed.

    The following information is gathered by the tool:
    * Tenant
    * Domain
    * Users
    * Licenses
    * ServicePrincipals
    * DirectoryRoles
    * DirectoryRoleMembers
    * Groups
    * GroupMembers
    * Devices

    Author     : Prashant Mahajan

.NOTES

    The following commands can be used to turn off ExecutionPolicy: (Requires Admin Privs)

    PS > $ExecPolicy = Get-ExecutionPolicy
    PS > Set-ExecutionPolicy bypass
    PS > .\AzureADRecon.ps1
    PS > Set-ExecutionPolicy $ExecPolicy

    OR

    Start the PowerShell as follows:
    powershell.exe -ep bypass

    OR

    Already have a PowerShell open ?
    PS > $Env:PSExecutionPolicyPreference = 'Bypass'

    OR

    powershell.exe -nologo -executionpolicy bypass -noprofile -file AzureADRecon.ps1

.PARAMETER Method
	Which method to use; AzureAD (default)

.PARAMETER Credential
	Domain Credentials.

.PARAMETER GenExcel
	Path for AzureADRecon output folder containing the CSV files to generate the AzureADRecon-Report.xlsx. Use it to generate the AzureADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run AzureADRecon.

.PARAMETER TenantID
    The Azure TenantID to connect to when you have multiple tenants.

.PARAMETER OutputDir
	Path for AzureADRecon output folder to save the files and the AzureADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)

.PARAMETER Collect
    Which modules to run; Comma separated; e.g Tenant,Domain (Default all)
    Valid values include: Tenant, Domain, Licenses, Users, ServicePrincipals, DirectoryRoles, DirectoryRoleMembers, Groups, GroupMembers, Devices.

.PARAMETER OutputType
    Output Type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

.PARAMETER Threads
    The number of threads to use during processing objects. (Default 10)

.PARAMETER Log
    Create AzureADRecon Log using Start-Transcript

.EXAMPLE

    .\AzureADRecon.ps1

    Or

    $username = "username@fqdn"
    $passwd = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
    $creds = New-Object System.Management.Automation.PSCredential ($username, $passwd)
	.\AzureADRecon.ps1 -Credential $creds

.EXAMPLE

	.\AzureADRecon.ps1 -GenExcel C:\AzureADRecon-Report-<timestamp>
    [*] AzureADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Generating AzureADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\AzureADRecon-Report-<timestamp>\AzureADRecon-Report.xlsx

.LINK

    https://github.com/adrecon/AzureADRecon
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false, HelpMessage = "Which method to use; AzureAD (default).")]
    [ValidateSet('AzureAD')]
    [string] $Method = 'AzureAD',

    [Parameter(Mandatory = $false, HelpMessage = "Which Tenant ID to connect to, picks up the default tenant ID if nothing specified")]
    [string] $TenantID,

    [Parameter(Mandatory = $false, HelpMessage = "Azure Credentials.")]
    [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false, HelpMessage = "Path for AzureADRecon output folder containing the CSV files to generate the AzureADRecon-Report.xlsx. Use it to generate the AzureADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run AzureADRecon.")]
    [string] $GenExcel,

    [Parameter(Mandatory = $false, HelpMessage = "Path for AzureADRecon output folder to save the CSV/XML/JSON/HTML files and the AzureADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)")]
    [string] $OutputDir,

    [Parameter(Mandatory = $false, HelpMessage = "Which modules to run; Comma separated; e.g Tenant,Domain (Default all) Valid values include: Tenant, Domain, Licenses, Users, ServicePrincipals, DirectoryRoles, DirectoryRoleMembers, Groups, GroupMembers, Devices")]
    [ValidateSet('Tenant', 'Domain', 'Licenses', 'Users', 'ServicePrincipals', 'DirectoryRoles', 'DirectoryRoleMembers', 'Groups', 'GroupMembers', 'Devices', 'Default')]
    [array] $Collect = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "Output type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel)")]
    [ValidateSet('STDOUT', 'CSV', 'XML', 'JSON', 'EXCEL', 'HTML', 'All', 'Default')]
    [array] $OutputType = 'Default',

    [Parameter(Mandatory = $false, HelpMessage = "The number of threads to use during processing of objects. Default 10")]
    [ValidateRange(1,100)]
    [int] $Threads = 10,

    [Parameter(Mandatory = $false, HelpMessage = "Create AzureADRecon Log using Start-Transcript")]
    [switch] $Log
)

$AzureADSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading;

namespace AADRecon
{
    public static class AzureADClass
    {
        //Values taken from https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference
        private static Dictionary<string, string> AzureSkuIDDictionary = new Dictionary<string, string>()
        {
            {"0c266dff-15dd-4b49-8397-2bb16070ed52", "Audio Conferencing"},
            {"2b9c8e7c-319c-43a2-a2a0-48c5c6161de7", "Azure Active Directory Basic"},
            {"078d2b04-f1bd-4111-bbd4-b4b1b354cef4", "Azure Active Directory Premium P1"},
            {"84a661c4-e949-4bd2-a560-ed7766fcaf2b", "Azure Active Directory Premium P2"},
            {"c52ea49f-fe5d-4e95-93ba-1de91d380f89", "Azure Information Protection Plan 1"},
            {"ea126fc5-a19e-42e2-a731-da9d437bffcf", "Dynamics 365 Customer Engagement Plan Enterprise Edition"},
            {"749742bf-0d37-4158-a120-33567104deeb", "Dynamics 365 For Customer Service Enterprise Edition"},
            {"cc13a803-544e-4464-b4e4-6d6169a138fa", "Dynamics 365 For Financials Business Edition"},
            {"8edc2cf8-6438-4fa9-b6e3-aa1660c640cc", "Dynamics 365 For Sales And Customer Service Enterprise Edition"},
            {"1e1a282c-9c54-43a2-9310-98ef728faace", "Dynamics 365 For Sales Enterprise Edition"},
            {"8e7a3d30-d97d-43ab-837c-d7701cef83dc", "Dynamics 365 For Team Members Enterprise Edition"},
            {"ccba3cfe-71ef-423a-bd87-b6df3dce59a9", "Dynamics 365 Unf Ops Plan Ent Edition"},
            {"efccb6f7-5641-4e0e-bd10-b4976e1bf68e", "Enterprise Mobility + Security E3"},
            {"b05e124f-c7cc-45a0-a6aa-8cf78c946968", "Enterprise Mobility + Security E5"},
            {"4b9405b0-7788-4568-add1-99614e613b69", "Exchange Online (Plan 1)"},
            {"19ec0d23-8335-4cbd-94ac-6050e30712fa", "Exchange Online (Plan 2)"},
            {"ee02fd1b-340e-4a4b-b355-4a514e4c8943", "Exchange Online Archiving For Exchange Online"},
            {"90b5e015-709a-4b8b-b08e-3200f994494c", "Exchange Online Archiving For Exchange Server"},
            {"7fc0182e-d107-4556-8329-7caaa511197b", "Exchange Online Essentials"},
            {"e8f81a67-bd96-4074-b108-cf193eb9433b", "Exchange Online Essentials"},
            {"80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82", "Exchange Online Kiosk"},
            {"cb0a98a8-11bc-494c-83d9-c1b1ac65327e", "Exchange Online Pop"},
            {"061f9ace-7d42-4136-88ac-31dc755f143f", "Intune"},
            {"b17653a4-2443-4e8c-a550-18249dda78bb", "Microsoft 365 A1"},
            {"4b590615-0888-425a-a965-b3bf7789848d", "Microsoft 365 A3 For Faculty"},
            {"7cfd9a2b-e110-4c39-bf20-c6a3f36a3121", "Microsoft 365 A3 For Students"},
            {"e97c048c-37a4-45fb-ab50-922fbf07a370", "Microsoft 365 A5 For Faculty"},
            {"46c119d4-0379-4a9d-85e4-97c66d3f909e", "Microsoft 365 A5 For Students"},
            {"cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46", "Microsoft 365 Business"},
            {"05e9a617-0261-4cee-bb44-138d3ef5d965", "Microsoft 365 E3"},
            {"06ebc4ee-1bb5-47dd-8120-11324bc54e06", "Microsoft 365 E5"},
            {"d61d61cc-f992-433f-a577-5bd016037eeb", "Microsoft 365 E3_Usgov_Dod"},
            {"ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658", "Microsoft 365 E3_Usgov_Gcchigh"},
            {"184efa21-98c3-4e5d-95ab-d07053a96e67", "Microsoft 365 E5 Compliance"},
            {"26124093-3d78-432b-b5dc-48bf992543d5", "Microsoft 365 E5 Security"},
            {"44ac31e7-2999-4304-ad94-c948886741d4", "Microsoft 365 E5 Security For Ems E5"},
            {"66b55226-6b4f-492c-910c-a3b7a3c9d993", "Microsoft 365 F1"},
            {"111046dd-295b-4d6d-9724-d52ac90bd1f2", "Microsoft Defender Advanced Threat Protection"},
            {"906af65a-2970-46d5-9b58-4e9aa50f0657", "Microsoft Dynamics Crm Online Basic"},
            {"d17b27af-3f49-4822-99f9-56a661538792", "Microsoft Dynamics Crm Online"},
            {"ba9a34de-4489-469d-879c-0f0f145321cd", "Ms Imagine Academy"},
            {"a4585165-0533-458a-97e3-c400570268c4", "Office 365 A5 For Faculty"},
            {"ee656612-49fa-43e5-b67e-cb1fdf7699df", "Office 365 A5 For Students"},
            {"1b1b1f7a-8355-43b6-829f-336cfccb744c", "Office 365 Advanced Compliance"},
            {"4ef96642-f096-40de-a3e9-d83fb2f90211", "Office 365 Advanced Threat Protection (Plan 1)"},
            {"cdd28e44-67e3-425e-be4c-737fab2899d3", "Office 365 Business"},
            {"b214fe43-f5a3-4703-beeb-fa97188220fc", "Office 365 Business"},
            {"3b555118-da6a-4418-894f-7df1e2096870", "Office 365 Business Essentials"},
            {"dab7782a-93b1-4074-8bb1-0e61318bea0b", "Office 365 Business Essentials"},
            {"f245ecc8-75af-4f8e-b61f-27d8114de5f3", "Office 365 Business Premium"},
            {"ac5cef5d-921b-4f97-9ef3-c99076e5470f", "Office 365 Business Premium"},
            {"18181a46-0d4e-45cd-891e-60aabd171b4e", "Office 365 E1"},
            {"6634e0ce-1a9f-428c-a498-f84ec7b8aa2e", "Office 365 E2"},
            {"6fd2c87f-b296-42f0-b197-1e91e994b900", "Office 365 E3"},
            {"189a915c-fe4f-4ffa-bde4-85b9628d07a0", "Office 365 E3 Developer"},
            {"b107e5a3-3e60-4c0d-a184-a7e4395eb44c", "Office 365 E3_Usgov_Dod"},
            {"aea38a85-9bd5-4981-aa00-616b411205bf", "Office 365 E3_Usgov_Gcchigh"},
            {"1392051d-0cb9-4b7a-88d5-621fee5e8711", "Office 365 E4"},
            {"c7df2760-2c81-4ef7-b578-5b5392b571df", "Office 365 E5"},
            {"26d45bd9-adf1-46cd-a9e1-51e9a5524128", "Office 365 E5 Without Audio Conferencing"},
            {"4b585984-651b-448a-9e53-3b10f069cf7f", "Office 365 F1"},
            {"04a7fb0d-32e0-4241-b4f5-3f7618cd1162", "Office 365 Midsize Business"},
            {"c2273bd0-dff7-4215-9ef5-2c7bcfb06425", "Office 365 Proplus"},
            {"bd09678e-b83c-4d3f-aaba-3dad4abd128b", "Office 365 Small Business"},
            {"fc14ec4a-4169-49a4-a51e-2c852931814b", "Office 365 Small Business Premium"},
            {"e6778190-713e-4e4f-9119-8b8238de25df", "Onedrive For Business (Plan 1)"},
            {"ed01faf2-1d88-4947-ae91-45ca18703a96", "Onedrive For Business (Plan 2)"},
            {"b30411f5-fea1-4a59-9ad9-3db7c7ead579", "Power Apps Per User Plan"},
            {"45bc2c81-6072-436a-9b0b-3b12eefbc402", "Power Bi For Office 365 Add-On"},
            {"f8a1db68-be16-40ed-86d5-cb42ce701560", "Power Bi Pro"},
            {"a10d5e58-74da-4312-95c8-76be4e5b75a0", "Project For Office 365"},
            {"776df282-9fc0-4862-99e2-70e561b9909e", "Project Online Essentials"},
            {"09015f9f-377f-4538-bbb5-f75ceb09358a", "Project Online Premium"},
            {"2db84718-652c-47a7-860c-f10d8abbdae3", "Project Online Premium Without Project Client"},
            {"53818b1b-4a27-454b-8896-0dba576410e6", "Project Online Professional"},
            {"f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c", "Project Online With Project For Office 365"},
            {"1fc08a02-8b3d-43b9-831e-f76859e04e1a", "Sharepoint Online (Plan 1)"},
            {"a9732ec9-17d9-494c-a51c-d6b45b384dcb", "Sharepoint Online (Plan 2)"},
            {"e43b5b99-8dfb-405f-9987-dc307f34bcbd", "Skype For Business Cloud Pbx"},
            {"b8b749f8-a4ef-4887-9539-c95b1eaa5db7", "Skype For Business Online (Plan 1)"},
            {"d42c793f-6c78-4f43-92ca-e8f6a02b035f", "Skype For Business Online (Plan 2)"},
            {"d3b4fe1f-9992-4930-8acb-ca6ec609365e", "Skype For Business Pstn Domestic And International Calling"},
            {"0dab259f-bf13-4952-b7f8-7db8f131b28d", "Skype For Business Pstn Domestic Calling"},
            {"54a152dc-90de-4996-93d2-bc47e670fc06", "Skype For Business Pstn Domestic Calling (120 Minutes)"},
            {"4b244418-9658-4451-a2b8-b5e2b364e9bd", "Visio Online Plan 1"},
            {"c5928f49-12ba-48f7-ada3-0d743a3601d5", "Visio Online Plan 2"},
            {"cb10e6cd-9da4-4992-867b-67546b1db821", "Windows 10 Enterprise E3"},
            {"488ba24a-39a9-4473-8ee5-19291e71b002", "Windows 10 Enterprise E5"}
        };

        // Add missing SkuIDs to the dictionary
        private static void UpdateSkuIDDictionary(Object[] AdLicenses)
        {
            foreach (PSObject AdLicense in AdLicenses)
            {
                if (!AzureADClass.AzureSkuIDDictionary.ContainsKey(Convert.ToString(AdLicense.Members["SkuId"].Value)))
                {
                    AzureADClass.AzureSkuIDDictionary.Add(Convert.ToString(AdLicense.Members["SkuId"].Value),Convert.ToString(AdLicense.Members["SkuPartNumber"].Value));
                }
            }
        }

		private static readonly Dictionary<string, string> Replacements = new Dictionary<string, string>()
        {
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static string CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            string CleanedString = string.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (string Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static int ObjectCount(Object[] ADRObject)
        {
            return ADRObject.Length;
        }

        public static Object[] TenantParser(Object[] AdTenant, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdTenant, numOfThreads, "Tenant");
            return ADRObj;
        }

        public static Object[] DomainParser(Object[] AdDomain, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDomain, numOfThreads, "Domain");
            return ADRObj;
        }

        public static Object[] LicenseParser(Object[] AdLicenses, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdLicenses, numOfThreads, "Licenses");
            return ADRObj;
        }

        public static Object[] UserParser(Object[] AdUsers, Object[] AdLicenses, int numOfThreads)
        {
            AzureADClass.UpdateSkuIDDictionary(AdLicenses);
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "Users");
            return ADRObj;
        }

        public static Object[] ServicePrincipalParser(Object[] AdServicePrincipals, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdServicePrincipals, numOfThreads, "ServicePrincipals");
            return ADRObj;
        }

        public static Object[] DirectoryRoleParser(Object[] AdDirectoryRoles, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDirectoryRoles, numOfThreads, "DirectoryRoles");
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "Groups");
            return ADRObj;
        }

        public static Object[] DeviceParser(Object[] AdDevices, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDevices, numOfThreads, "Devices");
            return ADRObj;
        }

        static Object[] runProcessor(Object[] arrayToProcess, int numOfThreads, string processorType)
        {
            int totalRecords = arrayToProcess.Length;
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = new SimpleResultsHandler ();
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            return resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(string name)
        {
            switch (name)
            {
                case "Tenant":
                    return new TenantRecordProcessor();
                case "Domain":
                    return new DomainRecordProcessor();
                case "Licenses":
                    return new LicenseRecordProcessor();
                case "Users":
                    return new UserRecordProcessor();
                case "ServicePrincipals":
                    return new ServicePrincipalRecordProcessor();
                case "DirectoryRoles":
                    return new DirectoryRoleRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "Devices":
                    return new DeviceRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            PSObject[] processRecord(Object record);
        }

        class TenantRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADTenant = (PSObject) record;

                    List<PSObject> AzureADTenantObjList = new List<PSObject>();
                    PSObject AzureADTenantObj = new PSObject();
                    int icount = 0;

                    List<Microsoft.Open.AzureAD.Model.VerifiedDomain> VerifiedDomainsList = new List<Microsoft.Open.AzureAD.Model.VerifiedDomain>();
                    List<string> TechnicalNotificationMailsList = new List<string>();
                    List<Microsoft.Open.AzureAD.Model.AssignedPlan> AssignedPlanList = new List<Microsoft.Open.AzureAD.Model.AssignedPlan>();
                    int count = 0;
                    string TechnicalNotificationMails = null;

                    Object[] ObjValues = new Object[]{
                        "DisplayName", CleanString(AzureADTenant.Members["DisplayName"].Value)
                    };

                    for (icount = 0; icount < ObjValues.Length; icount++)
                    {
                        AzureADTenantObj = new PSObject();
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                        AzureADTenantObjList.Add( AzureADTenantObj );
                        icount++;
                    }

                    if (((List<Microsoft.Open.AzureAD.Model.VerifiedDomain>) AzureADTenant.Members["VerifiedDomains"].Value).Count != 0)
                    {
                        VerifiedDomainsList = (List<Microsoft.Open.AzureAD.Model.VerifiedDomain>) AzureADTenant.Members["VerifiedDomains"].Value;
                        count = 0;
                        foreach (Microsoft.Open.AzureAD.Model.VerifiedDomain value in VerifiedDomainsList)
                        {
                            ObjValues = new Object[]{
                                "VerifiedDomain(" + count + ") - Name", value.Name,
                                "VerifiedDomain(" + count + ") - Type", value.Type,
                                "VerifiedDomain(" + count + ") - Capabilities", value.Capabilities,
                                "VerifiedDomain(" + count + ") - _Default", value._Default,
                                "VerifiedDomain(" + count + ") - Initial", value.Initial,
                                "VerifiedDomain(" + count + ") - Id", value.Id
                            };
                            for (icount = 0; icount < ObjValues.Length; icount++)
                            {
                                AzureADTenantObj = new PSObject();
                                AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                                AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                                AzureADTenantObjList.Add( AzureADTenantObj );
                                icount++;
                            }
                            count++;
                        }
                    }

                    if (((List<string>) AzureADTenant.Members["TechnicalNotificationMails"].Value).Count != 0)
                    {
                        TechnicalNotificationMailsList = (List<string>) AzureADTenant.Members["TechnicalNotificationMails"].Value;
                        foreach (string value in TechnicalNotificationMailsList)
                        {
                            TechnicalNotificationMails = TechnicalNotificationMails + "," + value;
                        }
                        TechnicalNotificationMails = TechnicalNotificationMails.TrimStart(',');
                    }

                    ObjValues = new Object[]{
                        "DirSyncEnabled", AzureADTenant.Members["DirSyncEnabled"].Value,
                        "CompanyLastDirSyncTime", AzureADTenant.Members["CompanyLastDirSyncTime"].Value,
                        "Street", AzureADTenant.Members["Street"].Value,
                        "City", AzureADTenant.Members["City"].Value,
                        "PostalCode", AzureADTenant.Members["PostalCode"].Value,
                        "State", AzureADTenant.Members["State"].Value,
                        "Country", AzureADTenant.Members["Country"].Value,
                        "CountryLetterCode", AzureADTenant.Members["CountryLetterCode"].Value,
                        "TechnicalNotificationMails", TechnicalNotificationMails
                    };
                    for (icount = 0; icount < ObjValues.Length; icount++)
                    {
                        AzureADTenantObj = new PSObject();
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                        AzureADTenantObjList.Add( AzureADTenantObj );
                        icount++;
                    }

                    if (((List<Microsoft.Open.AzureAD.Model.AssignedPlan>) AzureADTenant.Members["AssignedPlans"].Value).Count != 0)
                    {
                        AssignedPlanList = (List<Microsoft.Open.AzureAD.Model.AssignedPlan>) AzureADTenant.Members["AssignedPlans"].Value;
                        count = 0;
                        foreach (Microsoft.Open.AzureAD.Model.AssignedPlan value in AssignedPlanList)
                        {
                            ObjValues = new Object[]{
                                "AssignedPlan(" + count + ") - AssignedTimestamp", value.AssignedTimestamp,
                                "AssignedPlan(" + count + ") - CapabilityStatus", value.CapabilityStatus,
                                "AssignedPlan(" + count + ") - Service", value.Service,
                                "AssignedPlan(" + count + ") - ServicePlanId", value.ServicePlanId
                            };
                            for (icount = 0; icount < ObjValues.Length; icount++)
                            {
                                AzureADTenantObj = new PSObject();
                                AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                                AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                                AzureADTenantObjList.Add( AzureADTenantObj );
                                icount++;
                            }
                            count++;
                        }
                    }

                    ObjValues = new Object[]{
                        "ObjectType", AzureADTenant.Members["ObjectType"].Value,
                        "ObjectId", AzureADTenant.Members["ObjectId"].Value
                    };
                    for (icount = 0; icount < ObjValues.Length; icount++)
                    {
                        AzureADTenantObj = new PSObject();
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Category", ObjValues[icount]));
                        AzureADTenantObj.Members.Add(new PSNoteProperty("Value", ObjValues[icount+1]));
                        AzureADTenantObjList.Add( AzureADTenantObj );
                        icount++;
                    }

                    return AzureADTenantObjList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class DomainRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADDomain = (PSObject) record;

                    List<string> SupportedServicesList = new List<string>();
                    string SupportedServices = null;

                    PSObject AzureADDomainObj = new PSObject();
                    AzureADDomainObj.Members.Add(new PSNoteProperty("Name", CleanString(AzureADDomain.Members["Name"].Value)));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("AuthenticationType", AzureADDomain.Members["authenticationType"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("AvailabilityStatus", AzureADDomain.Members["availabilityStatus"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isAdminManaged", AzureADDomain.Members["isAdminManaged"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isDefault", AzureADDomain.Members["isDefault"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isInitial", AzureADDomain.Members["isInitial"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isRoot", AzureADDomain.Members["isRoot"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("isVerified", AzureADDomain.Members["isVerified"].Value));

                    if (((List<string>) AzureADDomain.Members["SupportedServices"].Value).Count != 0)
                    {
                        SupportedServicesList = (List<string>) AzureADDomain.Members["SupportedServices"].Value;
                        foreach (string value in SupportedServicesList)
                        {
                            SupportedServices = SupportedServices + "," + value;
                        }
                        SupportedServices = SupportedServices.TrimStart(',');
                    }
                    AzureADDomainObj.Members.Add(new PSNoteProperty("SupportedServices", SupportedServices));

                    AzureADDomainObj.Members.Add(new PSNoteProperty("ForceDeleteState", AzureADDomain.Members["ForceDeleteState"].Value));
                    AzureADDomainObj.Members.Add(new PSNoteProperty("State", AzureADDomain.Members["State"].Value));
                    return new PSObject[] { AzureADDomainObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class LicenseRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADLicense = (PSObject) record;

                    string ServicePlanName = null;
                    List<Microsoft.Open.AzureAD.Model.ServicePlanInfo> ServicePlansList = new List<Microsoft.Open.AzureAD.Model.ServicePlanInfo>();

                    PSObject AzureADLicenseObj = new PSObject();

                    AzureADLicenseObj.Members.Add(new PSNoteProperty("SkuPartNumber", AzureADLicense.Members["SkuPartNumber"].Value));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("SkuId", AzureADLicense.Members["SkuId"].Value));

                    AzureADLicenseObj.Members.Add(new PSNoteProperty("ConsumedUnits", AzureADLicense.Members["ConsumedUnits"].Value));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("AppliesTo", AzureADLicense.Members["AppliesTo"].Value));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("CapabilityStatus", AzureADLicense.Members["CapabilityStatus"].Value));

                    if (((List<Microsoft.Open.AzureAD.Model.ServicePlanInfo>) AzureADLicense.Members["ServicePlans"].Value).Count != 0)
                    {
                        ServicePlansList = (List<Microsoft.Open.AzureAD.Model.ServicePlanInfo>) AzureADLicense.Members["ServicePlans"].Value;
                        foreach (Microsoft.Open.AzureAD.Model.ServicePlanInfo value in ServicePlansList)
                        {
                            ServicePlanName = ServicePlanName + "," + Convert.ToString(value.ServicePlanName);
                        }
                        ServicePlanName = ServicePlanName.TrimStart(',');
                    }

                    AzureADLicenseObj.Members.Add(new PSNoteProperty("PrepaidUnits-Enabled", (((Microsoft.Open.AzureAD.Model.LicenseUnitsDetail) AzureADLicense.Members["PrepaidUnits"].Value).Enabled)));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("PrepaidUnits-Suspended", (((Microsoft.Open.AzureAD.Model.LicenseUnitsDetail) AzureADLicense.Members["PrepaidUnits"].Value).Suspended)));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("PrepaidUnits-Warning", (((Microsoft.Open.AzureAD.Model.LicenseUnitsDetail) AzureADLicense.Members["PrepaidUnits"].Value).Warning)));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("ServicePlans-Name", ServicePlanName));
                    AzureADLicenseObj.Members.Add(new PSNoteProperty("ObjectId", AzureADLicense.Members["ObjectId"].Value));

                    return new PSObject[] { AzureADLicenseObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureAdUser = (PSObject) record;

                    List<Microsoft.Open.AzureAD.Model.AssignedLicense> AssignedLicensesList = new List<Microsoft.Open.AzureAD.Model.AssignedLicense>();
                    string AssignedLicenses = null;
                    string AssignedLicensesName = null;

                    PSObject AzureADUserObj = new PSObject();
                    AzureADUserObj.Members.Add(new PSNoteProperty("UserPrincipalName", CleanString(AzureAdUser.Members["UserPrincipalName"].Value)));
                    AzureADUserObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AzureAdUser.Members["DisplayName"].Value)));
                    AzureADUserObj.Members.Add(new PSNoteProperty("Enabled", AzureAdUser.Members["AccountEnabled"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("UserType", AzureAdUser.Members["UserType"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("DirSyncEnabled", AzureAdUser.Members["DirSyncEnabled"].Value));

                    if (((List<Microsoft.Open.AzureAD.Model.AssignedLicense>) AzureAdUser.Members["AssignedLicenses"].Value).Count != 0)
                    {
                        AssignedLicensesList = (List<Microsoft.Open.AzureAD.Model.AssignedLicense>) AzureAdUser.Members["AssignedLicenses"].Value;
                        foreach (Microsoft.Open.AzureAD.Model.AssignedLicense value in AssignedLicensesList)
                        {
                            AssignedLicenses = AssignedLicenses + "," + Convert.ToString(value.SkuId);
                            AssignedLicensesName = AssignedLicensesName + "," + (AzureADClass.AzureSkuIDDictionary.ContainsKey(Convert.ToString(value.SkuId)) ? AzureADClass.AzureSkuIDDictionary[Convert.ToString(value.SkuId)] : Convert.ToString(value.SkuId));
                        }
                        AssignedLicenses = AssignedLicenses.TrimStart(',');
                        AssignedLicensesName = AssignedLicensesName.TrimStart(',');
                    }
                    AzureADUserObj.Members.Add(new PSNoteProperty("AssignedLicensesName", AssignedLicensesName));

                    AzureADUserObj.Members.Add(new PSNoteProperty("PasswordPolicies", AzureAdUser.Members["PasswordPolicies"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("OnPremisesSecurityIdentifier",  AzureAdUser.Members["OnPremisesSecurityIdentifier"] != null ? AzureAdUser.Members["OnPremisesSecurityIdentifier"].Value : null));
                    AzureADUserObj.Members.Add(new PSNoteProperty("OnPremisesDistinguishedName", CleanString(((Dictionary<string, string>) AzureAdUser.Members["ExtensionProperty"].Value)["onPremisesDistinguishedName"])));
                    AzureADUserObj.Members.Add(new PSNoteProperty("CreatedDateTime", ((Dictionary<string, string>) AzureAdUser.Members["ExtensionProperty"].Value)["createdDateTime"]));
                    AzureADUserObj.Members.Add(new PSNoteProperty("LastDirSyncTime", AzureAdUser.Members["LastDirSyncTime"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("EmployeeId", ((Dictionary<string, string>) AzureAdUser.Members["ExtensionProperty"].Value)["employeeId"]));
                    AzureADUserObj.Members.Add(new PSNoteProperty("Mobile", AzureAdUser.Members["Mobile"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("TelephoneNumber", AzureAdUser.Members["TelephoneNumber"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("userIdentities", AzureAdUser.Members["userIdentities"] != null ? AzureAdUser.Members["userIdentities"].Value : null));
                    AzureADUserObj.Members.Add(new PSNoteProperty("odata.type", ((Dictionary<string, string>) AzureAdUser.Members["ExtensionProperty"].Value)["odata.type"]));
                    AzureADUserObj.Members.Add(new PSNoteProperty("ObjectId", AzureAdUser.Members["ObjectId"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("ObjectType", AzureAdUser.Members["ObjectType"].Value));
                    AzureADUserObj.Members.Add(new PSNoteProperty("AssignedLicenses", AssignedLicenses));
                    return new PSObject[] { AzureADUserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class ServicePrincipalRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureAdServicePrincipal = (PSObject) record;

                    List<string> AADRList = new List<string>();
                    string AADRString = null;
                    int AADRCount = 0;
                    string AADRString2 = null;
                    List<Microsoft.Open.AzureAD.Model.AppRole> AppRolesList = new List<Microsoft.Open.AzureAD.Model.AppRole>();
                    List<Microsoft.Open.AzureAD.Model.OAuth2Permission> OAuth2PermissionList = new List<Microsoft.Open.AzureAD.Model.OAuth2Permission>();

                    PSObject AzureAdServicePrincipalObj = new PSObject();
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AzureAdServicePrincipal.Members["DisplayName"].Value)));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("Enabled", AzureAdServicePrincipal.Members["AccountEnabled"].Value));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("PublisherName", AzureAdServicePrincipal.Members["PublisherName"].Value));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("AppId", AzureAdServicePrincipal.Members["AppId"].Value));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("ObjectId", AzureAdServicePrincipal.Members["ObjectId"].Value));

                    AADRList = new List<string>();
                    AADRString = null;
                    if (((List<string>) AzureAdServicePrincipal.Members["ServicePrincipalNames"].Value).Count != 0)
                    {
                        AADRList = (List<string>) AzureAdServicePrincipal.Members["ServicePrincipalNames"].Value;
                        foreach (string value in AADRList)
                        {
                            AADRString = AADRString + "," + value;
                        }
                        AADRString = AADRString.TrimStart(',');
                    }
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("ServicePrincipalNames", AADRString));

                    AADRList = new List<string>();
                    AADRString = null;
                    if (((List<string>) AzureAdServicePrincipal.Members["ReplyUrls"].Value).Count != 0)
                    {
                        AADRList = (List<string>) AzureAdServicePrincipal.Members["ReplyUrls"].Value;
                        foreach (string value in AADRList)
                        {
                            AADRString = AADRString + "," + value;
                        }
                        AADRString = AADRString.TrimStart(',');
                    }
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("ReplyUrls", AADRString));

                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("LogoutUrl", AzureAdServicePrincipal.Members["LogoutUrl"].Value));

                    AADRCount = 0;
                    AADRString = null;
                    AADRString2 = null;
                    if (((List<Microsoft.Open.AzureAD.Model.AppRole>) AzureAdServicePrincipal.Members["AppRoles"].Value).Count != 0)
                    {
                        AppRolesList = (List<Microsoft.Open.AzureAD.Model.AppRole>) AzureAdServicePrincipal.Members["AppRoles"].Value;
                        AADRCount = AppRolesList.Count;
                        foreach (Microsoft.Open.AzureAD.Model.AppRole value in AppRolesList)
                        {
                            AADRList = new List<string>();
                            AADRString = null;
                            if (((List<string>) value.AllowedMemberTypes).Count != 0)
                            {
                                AADRList = (List<string>) value.AllowedMemberTypes;
                                foreach (string valuestring in AADRList)
                                {
                                    AADRString = AADRString + "," + valuestring;
                                }
                                AADRString = AADRString.TrimStart(',');
                            }
                            AADRString2 = AADRString2 + AADRString + "," + value.Description + "," + value.Id + "," + value.IsEnabled + "," + value.Value + ",";
                        }
                        AADRString2 = AADRString2.TrimStart(',');
                    }
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("AppRoles (Count)", AADRCount));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("AppRoles - AllowedMemberTypes,Description,Id,IsEnabled,Value", CleanString(AADRString2)));

                    AADRCount = 0;
                    AADRString = null;
                    if (((List<Microsoft.Open.AzureAD.Model.OAuth2Permission>) AzureAdServicePrincipal.Members["Oauth2Permissions"].Value).Count != 0)
                    {
                        OAuth2PermissionList = (List<Microsoft.Open.AzureAD.Model.OAuth2Permission>) AzureAdServicePrincipal.Members["Oauth2Permissions"].Value;
                        AADRCount = OAuth2PermissionList.Count;
                        foreach (Microsoft.Open.AzureAD.Model.OAuth2Permission value in OAuth2PermissionList)
                        {
                            AADRString = AADRString + "," + value.AdminConsentDescription + "," + value.AdminConsentDisplayName + "," + value.Id + "," + value.IsEnabled + "," + value.Type + "," + value.UserConsentDescription + "," + value.UserConsentDisplayName + ","  + value.Value + ",";
                        }
                        AADRString = AADRString.TrimStart(',');
                    }
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("OAuth2Permissions (Count)", AADRCount));
                    AzureAdServicePrincipalObj.Members.Add(new PSNoteProperty("OAuth2Permissions - AdminConsentDescription,AdminConsentDisplayName,Id,IsEnabled,Type,UserConsentDescription,UserConsentDisplayName,Value", CleanString(AADRString)));

                    return new PSObject[] { AzureAdServicePrincipalObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class DirectoryRoleRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADDirectoryRole = (PSObject) record;

                    PSObject AzureADDirectoryRoleObj = new PSObject();
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AzureADDirectoryRole.Members["DisplayName"].Value)));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("RoleDisabled", AzureADDirectoryRole.Members["RoleDisabled"].Value));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("IsSystem", AzureADDirectoryRole.Members["IsSystem"].Value));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("Description", CleanString(AzureADDirectoryRole.Members["Description"].Value)));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("RoleTemplateId", AzureADDirectoryRole.Members["RoleTemplateId"].Value));
                    AzureADDirectoryRoleObj.Members.Add(new PSNoteProperty("ObjectId", AzureADDirectoryRole.Members["ObjectId"].Value));
                    return new PSObject[] { AzureADDirectoryRoleObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADGroup = (PSObject) record;

                    PSObject AzureADGroupObj = new PSObject();
                    AzureADGroupObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AzureADGroup.Members["DisplayName"].Value)));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("DirSyncEnabled", AzureADGroup.Members["DirSyncEnabled"].Value));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("LastDirSyncTime", AzureADGroup.Members["LastDirSyncTime"].Value));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("OnPremisesSecurityIdentifier", AzureADGroup.Members["OnPremisesSecurityIdentifier"].Value));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("SecurityEnabled", AzureADGroup.Members["SecurityEnabled"].Value));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("Description", CleanString(AzureADGroup.Members["Description"].Value)));
                    AzureADGroupObj.Members.Add(new PSNoteProperty("ObjectId", AzureADGroup.Members["ObjectId"].Value));
                    return new PSObject[] { AzureADGroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        class DeviceRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AzureADDevice = (PSObject) record;

                    PSObject AzureADDeviceObj = new PSObject();
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AzureADDevice.Members["DisplayName"].Value)));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("AccountEnabled", CleanString(AzureADDevice.Members["AccountEnabled"].Value)));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DirSyncEnabled", AzureADDevice.Members["DirSyncEnabled"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("LastDirSyncTime", AzureADDevice.Members["LastDirSyncTime"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DeviceOSType", AzureADDevice.Members["DeviceOSType"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DeviceOSVersion", AzureADDevice.Members["DeviceOSVersion"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("ApproximateLastLogonTimeStamp", AzureADDevice.Members["ApproximateLastLogonTimeStamp"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DeviceTrustType", AzureADDevice.Members["DeviceTrustType"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("ProfileType", AzureADDevice.Members["ProfileType"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("DeviceId", AzureADDevice.Members["DeviceId"].Value));
                    AzureADDeviceObj.Members.Add(new PSNoteProperty("ObjectId", AzureADDevice.Members["ObjectId"].Value));
                    return new PSObject[] { AzureADDeviceObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    return new PSObject[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            Object[] finalise();
        }

        class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<Object> processed = new List<Object>();

            public SimpleResultsHandler()
            {
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        for (var i = 0; i < results.Length; i++)
                        {
                            processed.Add((PSObject)results[i]);
                        }
                    }
                }
            }

            public Object[] finalise()
            {
                return processed.ToArray();
            }
        }
	}
}
"@

Function Get-DateDiff
{
<#
.SYNOPSIS
    Get difference between two dates.

.DESCRIPTION
    Returns the difference between two dates.

.PARAMETER Date1
    [DateTime]
    Date

.PARAMETER Date2
    [DateTime]
    Date

.OUTPUTS
    [System.ValueType.TimeSpan]
    Returns the difference between the two dates.
#>
    param (
        [Parameter(Mandatory = $true)]
        [DateTime] $Date1,

        [Parameter(Mandatory = $true)]
        [DateTime] $Date2
    )

    If ($Date2 -gt $Date1)
    {
        $DDiff = $Date2 - $Date1
    }
    Else
    {
        $DDiff = $Date1 - $Date2
    }
    Return $DDiff
}

Function Export-ADRCSV
{
<#
.SYNOPSIS
    Exports Object to a CSV file.

.DESCRIPTION
    Exports Object to a CSV file using Export-CSV.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the CSV File.

.OUTPUTS
    CSV file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        $ADRObj | Export-Csv -Path $ADFileName -NoTypeInformation -Encoding Default
    }
    Catch
    {
        Write-Warning "[Export-ADRCSV] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADRXML
{
<#
.SYNOPSIS
    Exports Object to a XML file.

.DESCRIPTION
    Exports Object to a XML file using Export-Clixml.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the XML File.

.OUTPUTS
    XML file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        (ConvertTo-Xml -NoTypeInformation -InputObject $ADRObj).Save($ADFileName)
    }
    Catch
    {
        Write-Warning "[Export-ADRXML] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADRJSON
{
<#
.SYNOPSIS
    Exports Object to a JSON file.

.DESCRIPTION
    Exports Object to a JSON file using ConvertTo-Json.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the JSON File.

.OUTPUTS
    JSON file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        ConvertTo-JSON -InputObject $ADRObj | Out-File -FilePath $ADFileName
    }
    Catch
    {
        Write-Warning "[Export-ADRJSON] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADRHTML
{
<#
.SYNOPSIS
    Exports Object to a HTML file.

.DESCRIPTION
    Exports Object to a HTML file using ConvertTo-Html.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the HTML File.

.OUTPUTS
    HTML file.
#>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName,

        [Parameter(Mandatory = $false)]
        [String] $AADROutputDir = $null
    )

$Header = @"
<style type="text/css">
th {
	color:white;
	background-color:blue;
}
td, th {
	border:0px solid black;
	border-collapse:collapse;
	white-space:pre;
}
tr:nth-child(2n+1) {
    background-color: #dddddd;
}
tr:hover td {
    background-color: #c1d5f8;
}
table, tr, td, th {
	padding: 0px;
	margin: 0px;
	white-space:pre;
}
table {
	margin-left:1px;
}
</style>
"@
    Try
    {
        If ($ADFileName.Contains("Index"))
        {
            $HTMLPath  = -join($AADROutputDir,'\','HTML-Files')
            $HTMLPath = $((Convert-Path $HTMLPath).TrimEnd("\"))
            $HTMLFiles = Get-ChildItem -Path $HTMLPath -name
            $HTML = $HTMLFiles | ConvertTo-HTML -Title "AzureADRecon" -Property @{Label="Table of Contents";Expression={"<a href='$($_)'>$($_)</a>"}} -Head $Header

            Add-Type -AssemblyName System.Web
            [System.Web.HttpUtility]::HtmlDecode($HTML) | Out-File -FilePath $ADFileName
        }
        Else
        {
            If ($ADRObj -is [array])
            {
                $ADRObj | Select-Object * | ConvertTo-HTML -As Table -Head $Header | Out-File -FilePath $ADFileName
            }
            Else
            {
                ConvertTo-HTML -InputObject $ADRObj -As Table -Head $Header | Out-File -FilePath $ADFileName
            }
        }
    }
    Catch
    {
        Write-Warning "[Export-ADRHTML] Failed to export $($ADFileName)."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
}

Function Export-ADR
{
<#
.SYNOPSIS
    Helper function for all output types supported.

.DESCRIPTION
    Helper function for all output types supported.

.PARAMETER ADObjectDN
    [PSObject]
    ADRObj

.PARAMETER ADROutputDir
    [String]
    Path for AzureADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.

.PARAMETER ADRModuleName
    [String]
    Module Name.

.OUTPUTS
    STDOUT, CSV, XML, JSON and/or HTML file, etc.
#>
    param(
        [Parameter(Mandatory = $true)]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [String] $AADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $true)]
        [String] $ADRModuleName
    )

    Switch ($OutputType)
    {
        'STDOUT'
        {
            If ($ADRModuleName -ne "AboutAzureADRecon")
            {
                If ($ADRObj -is [array])
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Out-String -Stream
                }
                Else
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Format-List | Out-String -Stream
                }
            }
        }
        'CSV'
        {
            $ADFileName  = -join($AADROutputDir,'\','CSV-Files','\',$ADRModuleName,'.csv')
            Export-ADRCSV -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'XML'
        {
            $ADFileName  = -join($AADROutputDir,'\','XML-Files','\',$ADRModuleName,'.xml')
            Export-ADRXML -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'JSON'
        {
            $ADFileName  = -join($AADROutputDir,'\','JSON-Files','\',$ADRModuleName,'.json')
            Export-ADRJSON -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'HTML'
        {
            $ADFileName  = -join($AADROutputDir,'\','HTML-Files','\',$ADRModuleName,'.html')
            Export-ADRHTML -ADRObj $ADRObj -ADFileName $ADFileName -AADROutputDir $AADROutputDir
        }
    }
}

Function Get-ADRExcelComObj
{
<#
.SYNOPSIS
    Creates a ComObject to interact with Microsoft Excel.

.DESCRIPTION
    Creates a ComObject to interact with Microsoft Excel if installed, else warning is raised.

.OUTPUTS
    [System.__ComObject] and [System.MarshalByRefObject]
    Creates global variables $excel and $workbook.
#>

    #Check if Excel is installed.
    Try
    {
        # Suppress verbose output
        $SaveVerbosePreference = $script:VerbosePreference
        $script:VerbosePreference = 'SilentlyContinue'
        $global:excel = New-Object -ComObject excel.application
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
    }
    Catch
    {
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
        Write-Warning "[Get-ADRExcelComObj] Excel does not appear to be installed. Skipping generation of AzureADRecon-Report.xlsx. Use the -GenExcel parameter to generate the AzureADRecon-Report.xslx on a host with Microsoft Excel installed."
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        Return $null
    }
    $excel.Visible = $true
    $excel.Interactive = $false
    $global:workbook = $excel.Workbooks.Add()
    If ($workbook.Worksheets.Count -eq 3)
    {
        $workbook.WorkSheets.Item(3).Delete()
        $workbook.WorkSheets.Item(2).Delete()
    }
}

Function Get-ADRExcelComObjRelease
{
<#
.SYNOPSIS
    Releases the ComObject created to interact with Microsoft Excel.

.DESCRIPTION
    Releases the ComObject created to interact with Microsoft Excel.

.PARAMETER ComObjtoRelease
    ComObjtoRelease

.PARAMETER Final
    Final
#>
    param(
        [Parameter(Mandatory = $true)]
        $ComObjtoRelease,

        [Parameter(Mandatory = $false)]
        [bool] $Final = $false
    )
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.releasecomobject(v=vs.110).aspx
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.finalreleasecomobject(v=vs.110).aspx
    If ($Final)
    {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObjtoRelease) | Out-Null
    }
    Else
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObjtoRelease) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Function Get-ADRExcelWorkbook
{
<#
.SYNOPSIS
    Adds a WorkSheet to the Workbook.

.DESCRIPTION
    Adds a WorkSheet to the Workbook using the $workboook global variable and assigns it a name.

.PARAMETER name
    [string]
    Name of the WorkSheet.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )

    $workbook.Worksheets.Add() | Out-Null
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = $name

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelImport
{
<#
.SYNOPSIS
    Helper to import CSV to the current WorkSheet.

.DESCRIPTION
    Helper to import CSV to the current WorkSheet. Supports two methods.

.PARAMETER ADFileName
    [string]
    Filename of the CSV file to import.

.PARAMETER method
    [int]
    Method to use for the import.

.PARAMETER row
    [int]
    Row.

.PARAMETER column
    [int]
    Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ADFileName,

        [Parameter(Mandatory = $false)]
        [int] $Method = 1,

        [Parameter(Mandatory = $false)]
        [int] $row = 1,

        [Parameter(Mandatory = $false)]
        [int] $column = 1
    )

    $excel.ScreenUpdating = $false
    If ($Method -eq 1)
    {
        If (Test-Path $ADFileName)
        {
            $worksheet = $workbook.Worksheets.Item(1)
            $TxtConnector = ("TEXT;" + $ADFileName)
            $CellRef = $worksheet.Range("A1")
            #Build, use and remove the text file connector
            $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)

            #65001: Unicode (UTF-8)
            $worksheet.QueryTables.item($Connector.name).TextFilePlatform = 65001
            $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
            $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
            $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
            $worksheet.QueryTables.item($Connector.name).delete()

            Get-ADRExcelComObjRelease -ComObjtoRelease $CellRef
            Remove-Variable CellRef
            Get-ADRExcelComObjRelease -ComObjtoRelease $Connector
            Remove-Variable Connector

            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        }
        Remove-Variable ADFileName
    }
    Elseif ($Method -eq 2)
    {
        $worksheet = $workbook.Worksheets.Item(1)
        If (Test-Path $ADFileName)
        {
            $ADTemp = Import-Csv -Path $ADFileName
            $ADTemp | ForEach-Object {
                Foreach ($prop in $_.PSObject.Properties)
                {
                    $worksheet.Cells.Item($row, $column) = $prop.Name
                    $worksheet.Cells.Item($row, $column + 1) = $prop.Value
                    $row++
                }
            }
            Remove-Variable ADTemp
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }
        Else
        {
            $worksheet.Cells.Item($row, $column) = "Error!"
        }
        Remove-Variable ADFileName
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

# Thanks Anant Shrivastava for the suggestion of using Pivot Tables for generation of the Stats sheets.
Function Get-ADRExcelPivotTable
{
<#
.SYNOPSIS
    Helper to add Pivot Table to the current WorkSheet.

.DESCRIPTION
    Helper to add Pivot Table to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER PivotTableName
    [string]
    Pivot Table Name.

.PARAMETER PivotRows
    [array]
    Row names from Source Sheet.

.PARAMETER PivotColumns
    [array]
    Column names from Source Sheet.

.PARAMETER PivotFilters
    [array]
    Row/Column names from Source Sheet to use as filters.

.PARAMETER PivotValues
    [array]
    Row/Column names from Source Sheet to use for Values.

.PARAMETER PivotPercentage
    [array]
    Row/Column names from Source Sheet to use for Percentage.

.PARAMETER PivotLocation
    [array]
    Location of the Pivot Table in Row/Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $false)]
        [array] $PivotRows,

        [Parameter(Mandatory = $false)]
        [array] $PivotColumns,

        [Parameter(Mandatory = $false)]
        [array] $PivotFilters,

        [Parameter(Mandatory = $false)]
        [array] $PivotValues,

        [Parameter(Mandatory = $false)]
        [array] $PivotPercentage,

        [Parameter(Mandatory = $false)]
        [string] $PivotLocation = "R1C1"
    )

    $excel.ScreenUpdating = $false
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)
    $workbook.ShowPivotTableFieldList = $false

    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottablesourcetype-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottableversionlist-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfieldorientation-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/constants-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfiltertype-enumeration-excel

    # xlDatabase = 1 # this just means local sheet data
    # xlPivotTableVersion12 = 3 # Excel 2007
    $PivotFailed = $false
    Try
    {
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $SrcWorksheet.UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
    }
    Catch
    {
        $PivotFailed = $true
        Write-Verbose "[PivotCaches().Create] Failed"
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
    }
    If ( $PivotFailed -eq $true )
    {
        $rows = $SrcWorksheet.UsedRange.Rows.Count
        ElseIf ($SrcSheetName -eq "Users")
        {
            $PivotCols = "A1:C"
        }
        $UsedRange = $SrcWorksheet.Range($PivotCols+$rows)
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
        Remove-Variable rows
	    Remove-Variable PivotCols
        Remove-Variable UsedRange
    }
    Remove-Variable PivotFailed
    $PivotTable = $PivotCaches.CreatePivotTable($PivotLocation,$PivotTableName)
    # $workbook.ShowPivotTableFieldList = $true

    If ($PivotRows)
    {
        ForEach ($Row in $PivotRows)
        {
            $PivotField = $PivotTable.PivotFields($Row)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        }
    }

    If ($PivotColumns)
    {
        ForEach ($Col in $PivotColumns)
        {
            $PivotField = $PivotTable.PivotFields($Col)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
        }
    }

    If ($PivotFilters)
    {
        ForEach ($Fil in $PivotFilters)
        {
            $PivotField = $PivotTable.PivotFields($Fil)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
        }
    }

    If ($PivotValues)
    {
        ForEach ($Val in $PivotValues)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        }
    }

    If ($PivotPercentage)
    {
        ForEach ($Val in $PivotPercentage)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
            $PivotField.Calculation = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlPercentOfTotal
            $PivotTable.ShowValuesRow = $false
        }
    }

    # $PivotFields.Caption = ""
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotField
    Remove-Variable PivotField
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotTable
    Remove-Variable PivotTable
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotCaches
    Remove-Variable PivotCaches
    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet
}

Function Get-ADRExcelAttributeStats
{
<#
.SYNOPSIS
    Helper to add Attribute Stats to the current WorkSheet.

.DESCRIPTION
    Helper to add Attribute Stats to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER Title1
    [string]
    Title1.

.PARAMETER PivotTableName
    [string]
    PivotTableName.

.PARAMETER PivotRows
    [string]
    PivotRows.

.PARAMETER PivotValues
    [string]
    PivotValues.

.PARAMETER PivotPercentage
    [string]
    PivotPercentage.

.PARAMETER Title2
    [string]
    Title2.

.PARAMETER ObjAttributes
    [OrderedDictionary]
    Attributes.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $Title1,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $true)]
        [string] $PivotRows,

        [Parameter(Mandatory = $true)]
        [string] $PivotValues,

        [Parameter(Mandatory = $true)]
        [string] $PivotPercentage,

        [Parameter(Mandatory = $true)]
        [string] $Title2,

        [Parameter(Mandatory = $true)]
        [System.Object] $ObjAttributes
    )

    $excel.ScreenUpdating = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)

    $row = 1
    $column = 1
    $worksheet.Cells.Item($row, $column) = $Title1
    $worksheet.Cells.Item($row,$column).Style = "Heading 2"
    $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $MergeCells = $worksheet.Range("A1:C1")
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true
    Remove-Variable MergeCells

    Get-ADRExcelPivotTable -SrcSheetName $SrcSheetName -PivotTableName $PivotTableName -PivotRows @($PivotRows) -PivotValues @($PivotValues) -PivotPercentage @($PivotPercentage) -PivotLocation "R2C1"
    $excel.ScreenUpdating = $false

    $row = 2
    "Type","Count","Percentage" | ForEach-Object {
        $worksheet.Cells.Item($row, $column) = $_
        $worksheet.Cells.Item($row, $column).Font.Bold = $True
        $column++
    }

    $row = 3
    $column = 1
    For($row = 3; $row -le 6; $row++)
    {
        $temptext = [string] $worksheet.Cells.Item($row, $column).Text
        switch ($temptext.ToUpper())
        {
            "TRUE" { $worksheet.Cells.Item($row, $column) = "Enabled" }
            "FALSE" { $worksheet.Cells.Item($row, $column) = "Disabled" }
            "GRAND TOTAL" { $worksheet.Cells.Item($row, $column) = "Total" }
        }
    }

    If ($ObjAttributes)
    {
        $row = 1
        $column = 6
        $worksheet.Cells.Item($row, $column) = $Title2
        $worksheet.Cells.Item($row,$column).Style = "Heading 2"
        $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
        $MergeCells = $worksheet.Range("F1:L1")
        $MergeCells.Select() | Out-Null
        $MergeCells.MergeCells = $true
        Remove-Variable MergeCells

        $row++
        "Category","Enabled Count","Enabled Percentage","Disabled Count","Disabled Percentage","Total Count","Total Percentage" | ForEach-Object {
            $worksheet.Cells.Item($row, $column) = $_
            $worksheet.Cells.Item($row, $column).Font.Bold = $True
            $column++
        }
        $ExcelColumn = ($SrcWorksheet.Columns.Find("Enabled"))
        $EnabledColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
        $column = 6
        $i = 2

        $ObjAttributes.keys | ForEach-Object {
            $ExcelColumn = ($SrcWorksheet.Columns.Find($_))
            $ColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
            $row++
            $i++
            $worksheet.Cells.Item($row, $column).Formula = "='" + $SrcWorksheet.Name + "'!" + $ExcelColumn.Address($false,$false)
            If ($_ -eq "PasswordPolicies")
            {
                # Remove count of "None"
                $worksheet.Cells.Item($row, $column+1).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"TRUE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')' + "-COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"TRUE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + '"None"' + ')'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+1).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"TRUE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            $worksheet.Cells.Item($row, $column+2).Formula = '=IFERROR(G' + $i + '/VLOOKUP("Enabled",A3:B6,2,FALSE),0)'
            If ($_ -eq "PasswordPolicies")
            {
                # Remove count of "None"
                $worksheet.Cells.Item($row, $column+3).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"FALSE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')' + "-COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"FALSE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + '"None"' + ')'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+3).Formula = "=COUNTIFS('" + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ',"FALSE",' + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            $worksheet.Cells.Item($row, $column+4).Formula = '=IFERROR(I' + $i + '/VLOOKUP("Disabled",A3:B6,2,FALSE),0)'
            If ($_ -eq "AssignedLicenses")
            {
                # Remove count of FieldName
                $worksheet.Cells.Item($row, $column+5).Formula = "=COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')-1'
            }
            ElseIf ($_ -eq "PasswordPolicies")
            {
                # Remove count of "None" and FieldName
                $worksheet.Cells.Item($row, $column+5).Formula = "=COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')' + "-COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + '"None"' + ')-1'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+5).Formula = "=COUNTIF('" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            $worksheet.Cells.Item($row, $column+6).Formula = '=IFERROR(K' + $i + '/VLOOKUP("Total",A3:B6,2,FALSE),0)'
        }

        # http://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
        "H", "J" , "L" | ForEach-Object {
            $rng = $_ + $($row - $ObjAttributes.Count + 1) + ":" + $_ + $($row)
            $worksheet.Range($rng).NumberFormat = "0.00%"
        }
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelChart
{
<#
.SYNOPSIS
    Helper to add charts to the current WorkSheet.

.DESCRIPTION
    Helper to add charts to the current WorkSheet.

.PARAMETER ChartType
    [int]
    Chart Type.

.PARAMETER ChartLayout
    [int]
    Chart Layout.

.PARAMETER ChartTitle
    [string]
    Title of the Chart.

.PARAMETER RangetoCover
    WorkSheet Range to be covered by the Chart.

.PARAMETER ChartData
    Data for the Chart.

.PARAMETER StartRow
    Start row to calculate data for the Chart.

.PARAMETER StartColumn
    Start column to calculate data for the Chart.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ChartType,

        [Parameter(Mandatory = $true)]
        [int] $ChartLayout,

        [Parameter(Mandatory = $true)]
        [string] $ChartTitle,

        [Parameter(Mandatory = $true)]
        $RangetoCover,

        [Parameter(Mandatory = $false)]
        $ChartData = $null,

        [Parameter(Mandatory = $false)]
        $StartRow = $null,

        [Parameter(Mandatory = $false)]
        $StartColumn = $null
    )

    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $chart = $worksheet.Shapes.AddChart().Chart
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlcharttype-enumeration-excel
    $chart.chartType = [int]([Microsoft.Office.Interop.Excel.XLChartType]::$ChartType)
    $chart.ApplyLayout($ChartLayout)
    If ($null -eq $ChartData)
    {
        If ($null -eq $StartRow)
        {
            $start = $worksheet.Range("A1")
        }
        Else
        {
            $start = $worksheet.Range($StartRow)
        }
        # get the last cell
        $X = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        If ($null -eq $StartColumn)
        {
            $start = $worksheet.Range("B1")
        }
        Else
        {
            $start = $worksheet.Range($StartColumn)
        }
        # get the last cell
        $Y = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        $ChartData = $worksheet.Range($X,$Y)

        Get-ADRExcelComObjRelease -ComObjtoRelease $X
        Remove-Variable X
        Get-ADRExcelComObjRelease -ComObjtoRelease $Y
        Remove-Variable Y
        Get-ADRExcelComObjRelease -ComObjtoRelease $start
        Remove-Variable start
    }
    $chart.SetSourceData($ChartData)
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chartclass.plotby?redirectedfrom=MSDN&view=excel-pia#Microsoft_Office_Interop_Excel_ChartClass_PlotBy
    $chart.PlotBy = [Microsoft.Office.Interop.Excel.XlRowCol]::xlColumns
    $chart.seriesCollection(1).Select() | Out-Null
    $chart.SeriesCollection(1).ApplyDataLabels() | out-Null
    # modify the chart title
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = $ChartTitle
    # Reposition the Chart
    $temp = $worksheet.Range($RangetoCover)
    # $chart.parent.placement = 3
    $chart.parent.top = $temp.Top
    $chart.parent.left = $temp.Left
    $chart.parent.width = $temp.Width
    # $chart.Legend.Delete()
    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $chart
    Remove-Variable chart
    Get-ADRExcelComObjRelease -ComObjtoRelease $ChartData
    Remove-Variable ChartData
    Get-ADRExcelComObjRelease -ComObjtoRelease $temp
    Remove-Variable temp
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelSort
{
<#
.SYNOPSIS
    Sorts a WorkSheet in the active Workbook.

.DESCRIPTION
    Sorts a WorkSheet in the active Workbook.

.PARAMETER ColumnName
    [string]
    Name of the Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )

    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Activate();

    $ExcelColumn = ($worksheet.Columns.Find($ColumnName))
    If ($ExcelColumn)
    {
        If ($ExcelColumn.Text -ne $ColumnName)
        {
            $BeginAddress = $ExcelColumn.Address(0,0,1,1)
            $End = $False
            Do {
                #Write-Verbose "[Get-ADRExcelSort] $($ExcelColumn.Text) selected instead of $($ColumnName) in the $($worksheet.Name) worksheet."
                $ExcelColumn = ($worksheet.Columns.FindNext($ExcelColumn))
                $Address = $ExcelColumn.Address(0,0,1,1)
                If ( ($Address -eq $BeginAddress) -or ($ExcelColumn.Text -eq $ColumnName) )
                {
                    $End = $True
                }
            } Until ($End -eq $True)
        }
        If ($ExcelColumn.Text -eq $ColumnName)
        {
            # Sort by Column
            $workSheet.ListObjects.Item(1).Sort.SortFields.Clear()
            $workSheet.ListObjects.Item(1).Sort.SortFields.Add($ExcelColumn) | Out-Null
            $worksheet.ListObjects.Item(1).Sort.Apply()
        }
        Else
        {
            Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet."
        }
    }
    Else
    {
        Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet."
    }
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Export-ADRExcel
{
<#
.SYNOPSIS
    Automates the generation of the AzureADRecon report.

.DESCRIPTION
    Automates the generation of the AzureADRecon report. If specific files exist, they are imported into the AzureADRecon report.

.PARAMETER ExcelPath
    [string]
    Path for AzureADRecon output folder containing the CSV files to generate the AzureADRecon-Report.xlsx

.OUTPUTS
    Creates the AzureADRecon-Report.xlsx report in the folder.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExcelPath
    )

    $ExcelPath = $((Convert-Path $ExcelPath).TrimEnd("\"))
    $ReportPath = -join($ExcelPath,'\','CSV-Files')
    If (!(Test-Path $ReportPath))
    {
        Write-Warning "[Export-ADRExcel] Could not locate the CSV-Files directory ... Exiting"
        Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
        Return $null
    }
    Get-ADRExcelComObj
    If ($excel)
    {
        Write-Output "[*] Generating AzureADRecon-Report.xlsx"

        $ADFileName = -join($ReportPath,'\','AboutAzureADRecon.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $workbook.Worksheets.Item(1).Name = "About AzureADRecon"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(3,2) , "https://github.com/adrecon/AzureADRecon", "" , "", "github.com/adrecon/AzureADRecon") | Out-Null
            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\','Tenant.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Tenant"
            Get-ADRExcelImport -ADFileName $ADFileName
            $TenantObj = Import-CSV -Path $ADFileName
            Remove-Variable ADFileName
            $TenantName = -join($TenantObj[0].Value,"-")
            Remove-Variable TenantObj
        }

        $ADFileName = -join($ReportPath,'\','Domain.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Domain"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Licenses.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Licenses"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Devices.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Devices"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','GroupMembers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "GroupMembers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "Group"
        }

        $ADFileName = -join($ReportPath,'\','Groups.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Groups"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "DisplayName"
        }

        $ADFileName = -join($ReportPath,'\','DirectoryRoleMembers.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DirectoryRoleMembers"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "DirectoryRole"
        }

        $ADFileName = -join($ReportPath,'\','DirectoryRoles.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "DirectoryRoles"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','ServicePrincipals.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "ServicePrincipals"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\','Users.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "Users"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName "UserPrincipalName"

            $worksheet = $workbook.Worksheets.Item(1)

            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            $worksheet.Cells.Item(1,3).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,3).font.ColorIndex = 2
            # Set Filter to Enabled Accounts only
            $worksheet.UsedRange.Select() | Out-Null
            $excel.Selection.AutoFilter(3,$true) | Out-Null
            $worksheet.Cells.Item(1,1).Select() | Out-Null
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # User Stats
        $ADFileName = -join($ReportPath,'\','Users.csv')
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "User Stats"
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add("DirSyncEnabled",'"TRUE"')
            $ObjAttributes.Add("AssignedLicenses",'"*"')
            $ObjAttributes.Add("PasswordPolicies",'"*"')

            Get-ADRExcelAttributeStats -SrcSheetName "Users" -Title1 "User Accounts in AzureAD" -PivotTableName "User Accounts Status" -PivotRows "Enabled" -PivotValues "UserPrincipalName" -PivotPercentage "UserPrincipalName"  -Title2 "Status of User Accounts in AzureAD" -ObjAttributes $ObjAttributes

            Get-ADRExcelChart -ChartType "xlPie" -ChartLayout 3 -ChartTitle "User Accounts in AzureAD" -RangetoCover "A9:D21" -ChartData $workbook.Worksheets.Item(1).Range("A3:A4,B3:B4")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(8,1) , "" , "Users!A1", "", "Raw Data") | Out-Null

            Get-ADRExcelChart -ChartType "xlBarClustered" -ChartLayout 1 -ChartTitle "Status of User Accounts in AzureAD" -RangetoCover "F9:L21" -ChartData $workbook.Worksheets.Item(1).Range("F2:F5,G2:G5")
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(8,6) , "" , "Users!A1", "", "Raw Data") | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # Create Table of Contents
        Get-ADRExcelWorkbook -Name "Table of Contents"
        $worksheet = $workbook.Worksheets.Item(1)

        $excel.ScreenUpdating = $false
        # Image format and properties
        # $path = "C:\AzureADRecon_Logo.jpg"
        # $base64aadrecon = [convert]::ToBase64String((Get-Content $path -Encoding byte))

		$base64aadrecon = "/9j/4AAQSkZJRgABAQAAkACQAAD/4QCeRXhpZgAATU0AKgAAAAgABQESAAMAAAABAAEAAAEaAAUAAAABAAAASgEbAAUAAAABAAAAUgEoAAMAAAABAAIAAIdpAAQAAAABAAAAWgAAAAAAAACQAAAAAQAAAJAAAAABAAOShgAHAAAAEgAAAISgAgAEAAAAAQAAAaagAwAEAAAAAQAAAD4AAAAAQVNDSUkAAABTY3JlZW5zaG90/+EJIWh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS40LjAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIi8+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPD94cGFja2V0IGVuZD0idyI/PgD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/+IPrElDQ19QUk9GSUxFAAEBAAAPnGFwcGwCEAAAbW50clJHQiBYWVogB+QAAQAHABYAIwAjYWNzcEFQUEwAAAAAQVBQTAAAAAAAAAAAAAAAAAAAAAAAAPbWAAEAAAAA0y1hcHBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARZGVzYwAAAVAAAABiZHNjbQAAAbQAAASCY3BydAAABjgAAAAjd3RwdAAABlwAAAAUclhZWgAABnAAAAAUZ1hZWgAABoQAAAAUYlhZWgAABpgAAAAUclRSQwAABqwAAAgMYWFyZwAADrgAAAAgdmNndAAADtgAAAAwbmRpbgAADwgAAAA+Y2hhZAAAD0gAAAAsbW1vZAAAD3QAAAAoYlRSQwAABqwAAAgMZ1RSQwAABqwAAAgMYWFiZwAADrgAAAAgYWFnZwAADrgAAAAgZGVzYwAAAAAAAAAIRGlzcGxheQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAG1sdWMAAAAAAAAAJgAAAAxockhSAAAAFAAAAdhrb0tSAAAADAAAAexuYk5PAAAAEgAAAfhpZAAAAAAAEgAAAgpodUhVAAAAFAAAAhxjc0NaAAAAFgAAAjBkYURLAAAAHAAAAkZubE5MAAAAFgAAAmJmaUZJAAAAEAAAAnhpdElUAAAAFAAAAohlc0VTAAAAEgAAApxyb1JPAAAAEgAAApxmckNBAAAAFgAAAq5hcgAAAAAAFAAAAsR1a1VBAAAAHAAAAthoZUlMAAAAFgAAAvR6aFRXAAAACgAAAwp2aVZOAAAADgAAAxRza1NLAAAAFgAAAyJ6aENOAAAACgAAAwpydVJVAAAAJAAAAzhlbkdCAAAAFAAAA1xmckZSAAAAFgAAA3BtcwAAAAAAEgAAA4ZoaUlOAAAAEgAAA5h0aFRIAAAADAAAA6pjYUVTAAAAGAAAA7ZlbkFVAAAAFAAAA1xlc1hMAAAAEgAAApxkZURFAAAAEAAAA85lblVTAAAAEgAAA95wdEJSAAAAGAAAA/BwbFBMAAAAEgAABAhlbEdSAAAAIgAABBpzdlNFAAAAEAAABDx0clRSAAAAFAAABExwdFBUAAAAFgAABGBqYUpQAAAADAAABHYATABDAEQAIAB1ACAAYgBvAGoAac7st+wAIABMAEMARABGAGEAcgBnAGUALQBMAEMARABMAEMARAAgAFcAYQByAG4AYQBTAHoA7QBuAGUAcwAgAEwAQwBEAEIAYQByAGUAdgBuAP0AIABMAEMARABMAEMARAAtAGYAYQByAHYAZQBzAGsA5gByAG0ASwBsAGUAdQByAGUAbgAtAEwAQwBEAFYA5AByAGkALQBMAEMARABMAEMARAAgAGMAbwBsAG8AcgBpAEwAQwBEACAAYwBvAGwAbwByAEEAQwBMACAAYwBvAHUAbABlAHUAciAPAEwAQwBEACAGRQZEBkgGRgYpBBoEPgQ7BEwEPgRABD4EMgQ4BDkAIABMAEMARCAPAEwAQwBEACAF5gXRBeIF1QXgBdlfaYJyAEwAQwBEAEwAQwBEACAATQDgAHUARgBhAHIAZQBiAG4A/QAgAEwAQwBEBCYEMgQ1BEIEPQQ+BDkAIAQWBBoALQQ0BDgEQQQ/BDsENQQ5AEMAbwBsAG8AdQByACAATABDAEQATABDAEQAIABjAG8AdQBsAGUAdQByAFcAYQByAG4AYQAgAEwAQwBECTAJAgkXCUAJKAAgAEwAQwBEAEwAQwBEACAOKg41AEwAQwBEACAAZQBuACAAYwBvAGwAbwByAEYAYQByAGIALQBMAEMARABDAG8AbABvAHIAIABMAEMARABMAEMARAAgAEMAbwBsAG8AcgBpAGQAbwBLAG8AbABvAHIAIABMAEMARAOIA7MDxwPBA8kDvAO3ACADvwO4A8wDvQO3ACAATABDAEQARgDkAHIAZwAtAEwAQwBEAFIAZQBuAGsAbABpACAATABDAEQATABDAEQAIABhACAAQwBvAHIAZQBzMKsw6TD8AEwAQwBEAAB0ZXh0AAAAAENvcHlyaWdodCBBcHBsZSBJbmMuLCAyMDIwAABYWVogAAAAAAAA8xYAAQAAAAEWylhZWiAAAAAAAACC9AAAPWT///+8WFlaIAAAAAAAAEwkAAC0hQAACuZYWVogAAAAAAAAJ74AAA4XAADIi2N1cnYAAAAAAAAEAAAAAAUACgAPABQAGQAeACMAKAAtADIANgA7AEAARQBKAE8AVABZAF4AYwBoAG0AcgB3AHwAgQCGAIsAkACVAJoAnwCjAKgArQCyALcAvADBAMYAywDQANUA2wDgAOUA6wDwAPYA+wEBAQcBDQETARkBHwElASsBMgE4AT4BRQFMAVIBWQFgAWcBbgF1AXwBgwGLAZIBmgGhAakBsQG5AcEByQHRAdkB4QHpAfIB+gIDAgwCFAIdAiYC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"

        $bytes = [System.Convert]::FromBase64String($base64aadrecon)
        Remove-Variable base64aadrecon

        $CompanyLogo = -join($ReportPath,'\','AzureADRecon_Logo.jpg')
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
        Add-Type -AssemblyName System.Drawing
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($CompanyLogo)

        Remove-Variable bytes
        Remove-Variable p
        Remove-Variable picture

        $LinkToFile = $false
        $SaveWithDocument = $true
        $Left = 0
        $Top = 0
        $Width = 180
        $Height = 50

        # Add image to the Sheet
        $worksheet.Shapes.AddPicture($CompanyLogo, $LinkToFile, $SaveWithDocument, $Left, $Top, $Width, $Height) | Out-Null

        Remove-Variable LinkToFile
        Remove-Variable SaveWithDocument
        Remove-Variable Left
        Remove-Variable Top
        Remove-Variable Width
        Remove-Variable Height

        If (Test-Path -Path $CompanyLogo)
        {
            Remove-Item $CompanyLogo
        }
        Remove-Variable CompanyLogo

        $row = 5
        $column = 1
        $worksheet.Cells.Item($row,$column)= "Table of Contents"
        $worksheet.Cells.Item($row,$column).Style = "Heading 2"
        $row++

        For($i=2; $i -le $workbook.Worksheets.Count; $i++)
        {
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,$column) , "" , "'$($workbook.Worksheets.Item($i).Name)'!A1", "", $workbook.Worksheets.Item($i).Name) | Out-Null
            $row++
        }

        $row++
		$workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,1) , "https://github.com/adrecon/AzureADRecon", "" , "", "github.com/adrecon/AzureADRecon") | Out-Null

        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        $excel.Windows.Item(1).Displaygridlines = $false
        $excel.ScreenUpdating = $true
        $ADStatFileName = -join($ExcelPath,'\',$TenantName,'AzureADRecon-Report.xlsx')
        Try
        {
            # Disable prompt if file exists
            $excel.DisplayAlerts = $False
            $workbook.SaveAs($ADStatFileName)
            Write-Output "[+] Excelsheet Saved to: $ADStatFileName"
        }
        Catch
        {
            Write-Error "[EXCEPTION] $($_.Exception.Message)"
        }
        $excel.Quit()
        Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet -Final $true
        Remove-Variable worksheet
        Get-ADRExcelComObjRelease -ComObjtoRelease $workbook -Final $true
        Remove-Variable -Name workbook -Scope Global
        Get-ADRExcelComObjRelease -ComObjtoRelease $excel -Final $true
        Remove-Variable -Name excel -Scope Global
    }
}

Function Get-AADRTenant
{
<#
.SYNOPSIS
    Returns information of the current AzureAD Tenant.

.DESCRIPTION
    Returns information of the current AzureAD Tenant.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRTenant = @( Get-AzureADTenantDetail -All $true )
        If ($AADRTenant)
        {
            $AADRTenantObj = [AADRecon.AzureADClass]::TenantParser($AADRTenant, $Threads)
        }
    }

    If ($AADRTenantObj)
    {
        Return $AADRTenantObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRDomain
{
<#
.SYNOPSIS
    Returns information of the current (or specified) AzureAD.

.DESCRIPTION
    Returns information of the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRDomain = @( Get-AzureADDomain )
        If ($AADRDomain)
        {
            $AADRDomainObj = [AADRecon.AzureADClass]::DomainParser($AADRDomain, $Threads)
        }
    }

    If ($AADRDomainObj)
    {
        Return $AADRDomainObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRLicense
{
<#
.SYNOPSIS
    Returns information of the current (or specified) AzureAD.

.DESCRIPTION
    Returns information of the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AzureADLicenses = @( Get-AzureADSubscribedSku )
        If ($AzureADLicenses)
        {
            $LicenseObj = [AADRecon.AzureADClass]::LicenseParser($AzureADLicenses, $Threads)
        }
    }

    If ($LicenseObj)
    {
        Return $LicenseObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRUser
{
<#
.SYNOPSIS
    Returns all users in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all users in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER date
    [DateTime]
    Date when AzureADRecon was executed.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AzureADUsers = @( Get-AzureADUser -All $true )
        $AzureADLicenses = @( Get-AzureADSubscribedSku )
        If($AzureADUsers)
        {
            Write-Verbose "[*] Total Users: $([AADRecon.AzureADClass]::ObjectCount($AzureADUsers))"
            $UserObj = [AADRecon.AzureADClass]::UserParser($AzureADUsers, $AzureADLicenses, $Threads)
        }
    }

    If ($UserObj)
    {
        Return $UserObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRServicePrincipal
{
<#
.SYNOPSIS
    Returns all service principals in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all service principals in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER date
    [DateTime]
    Date when AzureADRecon was executed.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AzureADServicePrincipals = @( Get-AzureADServicePrincipal -All $true )
        If($AzureADServicePrincipals)
        {
            Write-Verbose "[*] Total Service Principals: $([AADRecon.AzureADClass]::ObjectCount($AzureADServicePrincipals))"
            $ServicePrincipalObj = [AADRecon.AzureADClass]::ServicePrincipalParser($AzureADServicePrincipals, $Threads)
        }
    }

    If ($ServicePrincipalObj)
    {
        Return $ServicePrincipalObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRDirectoryRole
{
<#
.SYNOPSIS
    Returns all directory roles in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all directory roles in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRDirectoryRoles = @( Get-AzureADDirectoryRole )
        If ($AADRDirectoryRoles)
        {
            $AADRDirectoryRolesObj = [AADRecon.AzureADClass]::DirectoryRoleParser($AADRDirectoryRoles, $Threads)
        }
    }

    If ($AADRDirectoryRolesObj)
    {
        Return $AADRDirectoryRolesObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRDirectoryRoleMember
{
<#
.SYNOPSIS
    Returns all directory role membership in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all directory role membership in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRDirectoryRoles = @( Get-AzureADDirectoryRole )
        If ($AADRDirectoryRoles)
        {
            $AADRDirectoryRoleMemberObj = @()
            $AADRDirectoryRoles | ForEach-Object {
                $AADRDirectoryRoleMember = Get-AzureADDirectoryRoleMember -ObjectId $_.ObjectId
                $DirectoryRoleName = $([AADRecon.AzureADClass]::CleanString($_.DisplayName))
                If ($AADRDirectoryRoleMember)
                {
                    $AADRDirectoryRoleMember | ForEach-Object {
                        # Create the object for each instance.
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name DirectoryRole -Value $DirectoryRoleName
                        $Obj | Add-Member -MemberType NoteProperty -Name MemberName -Value $([AADRecon.AzureADClass]::CleanString($_.DisplayName))
                        $Obj | Add-Member -MemberType NoteProperty -Name MemberUserPrincipalName -Value $([AADRecon.AzureADClass]::CleanString($_.UserPrincipalName))
                        $AADRDirectoryRoleMemberObj += $Obj
                    }
                }
            }
        }
    }

    If ($AADRDirectoryRoleMemberObj)
    {
        Return $AADRDirectoryRoleMemberObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRGroup
{
<#
.SYNOPSIS
    Returns all groups in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all groups in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER date
    [DateTime]
    Date when AzureADRecon was executed.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRGroups = @( Get-AzureADGroup -All $true )
        If($AADRGroups)
        {
            Write-Verbose "[*] Total Groups: $([AADRecon.AzureADClass]::ObjectCount($AADRGroups))"
            $AADRGroupsObj = [AADRecon.AzureADClass]::GroupParser($AADRGroups, $Threads)
        }
    }

    If ($AADRGroupsObj)
    {
        Return $AADRGroupsObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRGroupMember
{
<#
.SYNOPSIS
    Returns all group membership in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all group membership in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRGroups = @( Get-AzureADGroup -All $true )
        If ($AADRGroups)
        {
            $AADRGroupMemberObj = @()
            $AADRGroups | ForEach-Object {
                $AADRGroupMemberMember = Get-AzureADGroupMember -ObjectId $_.ObjectId
                $GroupName = $([AADRecon.AzureADClass]::CleanString($_.DisplayName))
                If ($AADRGroupMemberMember)
                {
                    $AADRGroupMemberMember | ForEach-Object {
                        # Create the object for each instance.
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name GroupName -Value $GroupName
                        $Obj | Add-Member -MemberType NoteProperty -Name MemberName -Value $([AADRecon.AzureADClass]::CleanString($_.DisplayName))
                        $Obj | Add-Member -MemberType NoteProperty -Name MemberUserPrincipalName -Value $([AADRecon.AzureADClass]::CleanString($_.UserPrincipalName))
                        $AADRGroupMemberObj += $Obj
                    }
                }
            }
        }
    }

    If ($AADRGroupMemberObj)
    {
        Return $AADRGroupMemberObj
    }
    Else
    {
        Return $null
    }
}

Function Get-AADRDevice
{
<#
.SYNOPSIS
    Returns all devices in the current (or specified) AzureAD.

.DESCRIPTION
    Returns all devices in the current (or specified) AzureAD.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq 'AzureAD')
    {
        $AADRDevices = @( Get-AzureADDevice -All $true )
        If($AADRDevices)
        {
            Write-Verbose "[*] Total Devices: $([AADRecon.AzureADClass]::ObjectCount($AADRDevices))"
            $AADRDevicesObj = [AADRecon.AzureADClass]::DeviceParser($AADRDevices, $Threads)
        }
    }

    If ($AADRDevicesObj)
    {
        Return $AADRDevicesObj
    }
    Else
    {
        Return $null
    }
}

Function Remove-EmptyAADROutputDir
{
<#
.SYNOPSIS
    Removes AzureADRecon output folder if empty.

.DESCRIPTION
    Removes AzureADRecon output folder if empty.

.PARAMETER AADROutputDir
    [string]
	Path for AzureADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $AADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Switch ($OutputType)
    {
        'CSV'
        {
            $CSVPath  = -join($AADROutputDir,'\','CSV-Files')
            If (!(Test-Path -Path $CSVPath\*))
            {
                Write-Verbose "Removed Empty Directory $CSVPath"
                Remove-Item $CSVPath
            }
        }
        'XML'
        {
            $XMLPath  = -join($AADROutputDir,'\','XML-Files')
            If (!(Test-Path -Path $XMLPath\*))
            {
                Write-Verbose "Removed Empty Directory $XMLPath"
                Remove-Item $XMLPath
            }
        }
        'JSON'
        {
            $JSONPath  = -join($AADROutputDir,'\','JSON-Files')
            If (!(Test-Path -Path $JSONPath\*))
            {
                Write-Verbose "Removed Empty Directory $JSONPath"
                Remove-Item $JSONPath
            }
        }
        'HTML'
        {
            $HTMLPath  = -join($AADROutputDir,'\','HTML-Files')
            If (!(Test-Path -Path $HTMLPath\*))
            {
                Write-Verbose "Removed Empty Directory $HTMLPath"
                Remove-Item $HTMLPath
            }
        }
    }
    If (!(Test-Path -Path $AADROutputDir\*))
    {
        Remove-Item $AADROutputDir
        Write-Verbose "Removed Empty Directory $AADROutputDir"
    }
}

Function Get-AADRAbout
{
<#
.SYNOPSIS
    Returns information about AzureADRecon.

.DESCRIPTION
    Returns information about AzureADRecon.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER date
    [DateTime]
    Date

.PARAMETER AADReconVersion
    [string]
    AADRecon Version.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER RanonComputer
    [string]
    Details of the Computer running AzureADRecon.

.PARAMETER TotalTime
    [string]
    TotalTime.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $true)]
        [string] $AzureADReconVersion,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [string] $RanonComputer,

        [Parameter(Mandatory = $true)]
        [string] $TotalTime
    )

    $AboutAzureADRecon = @()

    $Version = $Method + " Version"

    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
    {
        $Username = $($Credential.UserName)
    }
    Else
    {
        $Username = $([Environment]::UserName)
    }

    $ObjValues = @("Date", $($date), "AzureADRecon", "https://github.com/adrecon/AzureADRecon", $Version, $($AzureADReconVersion), "Ran as user", $Username, "Ran on computer", $RanonComputer, "Execution Time (mins)", $($TotalTime))

    For ($i = 0; $i -lt $($ObjValues.Count); $i++)
    {
        $Obj = New-Object PSObject
        $Obj | Add-Member -MemberType NoteProperty -Name "Category" -Value $ObjValues[$i]
        $Obj | Add-Member -MemberType NoteProperty -Name "Value" -Value $ObjValues[$i+1]
        $i++
        $AboutAzureADRecon += $Obj
    }
    Return $AboutAzureADRecon
}

Function Invoke-AzureADRecon
{
<#
.SYNOPSIS
    Wrapper function to run AzureADRecon modules.

.DESCRIPTION
    Wrapper function to set variables, check dependencies and run AzureADRecon modules.

.PARAMETER Method
    [string]
    Which method to use; AzureAD.

.PARAMETER Collect
    [array]
    Which modules to run; Tenant, Domain, DirectoryRoles, Users, Groups.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER OutputDir
    [string]
	Path for AzureADRecon output folder to save the CSV files and the AzureADRecon-Report.xlsx.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    STDOUT, CSV, XML, JSON, HTML and/or Excel file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = $false)]
        [string] $GenExcel,

        [Parameter(Mandatory = $false)]
        [ValidateSet('AzureAD')]
        [string] $Method = 'AzureAD',

        [Parameter(Mandatory = $true)]
        [array] $Collect,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $false)]
        [string] $AADROutputDir,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    [string] $AzureADReconVersion = "v0.01"
    Write-Output "[*] AzureADRecon $AzureADReconVersion by Prashant Mahajan (@prashant3535)"

    If ($GenExcel)
    {
        If (!(Test-Path $GenExcel))
        {
            Write-Output "[Invoke-AzureADRecon] Invalid Path ... Exiting"
            Return $null
        }
        Export-ADRExcel -ExcelPath $GenExcel
        Return $null
    }

    # Suppress verbose output
    $SaveVerbosePreference = $script:VerbosePreference
    $script:VerbosePreference = 'SilentlyContinue'
    Try
    {
        If ($PSVersionTable.PSVersion.Major -ne 2)
        {
            $computer = Get-CimInstance -ClassName Win32_ComputerSystem
            $computerdomainrole = ($computer).DomainRole
        }
        Else
        {
            $computer = Get-WMIObject win32_computersystem
            $computerdomainrole = ($computer).DomainRole
        }
    }
    Catch
    {
        Write-Output "[Invoke-AzureADRecon] $($_.Exception.Message)"
    }
    If ($SaveVerbosePreference)
    {
        $script:VerbosePreference = $SaveVerbosePreference
        Remove-Variable SaveVerbosePreference
    }

    switch ($computerdomainrole)
    {
        0
        {
            [string] $computerrole = "Standalone Workstation"
            $Env:ADPS_LoadDefaultDrive = 0
            $UseAltCreds = $true
        }
        1 { [string] $computerrole = "Member Workstation" }
        2
        {
            [string] $computerrole = "Standalone Server"
            $UseAltCreds = $true
            $Env:ADPS_LoadDefaultDrive = 0
        }
        3 { [string] $computerrole = "Member Server" }
        4 { [string] $computerrole = "Backup Domain Controller" }
        5 { [string] $computerrole = "Primary Domain Controller" }
        default { Write-Output "Computer Role could not be identified." }
    }

    $RanonComputer = "$($computer.domain)\$([Environment]::MachineName) - $($computerrole)"
    Remove-Variable computer
    Remove-Variable computerdomainrole
    Remove-Variable computerrole

    # Import AzureAD module
    If ($Method -eq 'AzureAD')
    {
        If (Get-Module -ListAvailable -Name AzureAD)
        {
            Try
            {
                # Suppress verbose output on module import
                $SaveVerbosePreference = $script:VerbosePreference;
                $script:VerbosePreference = 'SilentlyContinue';
                Import-Module AzureAD -WarningAction Stop -ErrorAction Stop | Out-Null
                If ($SaveVerbosePreference)
                {
                    $script:VerbosePreference = $SaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
            }
            Catch
            {
                Write-Warning "[Invoke-AzureADRecon] Error importing AzureAD Module. Exiting"
                If ($SaveVerbosePreference)
                {
                    $script:VerbosePreference = $SaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
                Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                Return $null
            }
        }
        Else
        {
            Write-Warning "[Invoke-AzureADRecon] AzureAD Module is not installed. Run `Install-Module -Name AzureAD` to continue"
        }
    }

    # Compile C# code
    # Suppress Debug output
    $SaveDebugPreference = $script:DebugPreference
    $script:DebugPreference = 'SilentlyContinue'
    Try
    {
        If ($Method -eq 'AzureAD')
        {
            $AzureADModulePath = (Get-Module -ListAvailable AzureAD).path | Split-Path
            $CLR = ([System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion)[1]
            If ($CLR -eq "4")
            {
                Add-Type -TypeDefinition $($AzureADSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.ComponentModel.DataAnnotations")).Location
                    ($AzureADModulePath + "\Microsoft.Open.AzureAD16.Graph.Client.dll")
                ))
            }
            Else
            {
                Add-Type -TypeDefinition $($AzureADSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.ComponentModel.DataAnnotations")).Location
                    ($AzureADModulePath + "\Microsoft.Open.AzureAD16.Graph.Client.dll")
                ))
            }
            Remove-Variable AzureADModulePath
            Remove-Variable CLR
        }
    }
    Catch
    {
        Write-Output "[Invoke-AzureADRecon] $($_.Exception.Message)"
        Return $null
    }
    If ($SaveDebugPreference)
    {
        $script:DebugPreference = $SaveDebugPreference
        Remove-Variable SaveDebugPreference
    }

    Write-Output "[*] Running on $RanonComputer"

    Switch ($Collect)
    {
        'Tenant' { $AADRTenant = $true }
        'Domain' {$AADRDomain = $true }
        'Licenses' { $AADRLicenses = $true }
        'Users' { $AADRUsers = $true }
        'ServicePrincipals'{ $AADRServicePrincipals = $true }
        'DirectoryRoles' { $AADRDirectoryRoles = $true }
        'DirectoryRoleMembers' { $AADRDirectoryRoleMembers = $true }
        'Groups' {$AADRGroups = $true }
        'GroupMembers' {$AADRGroupMembers = $true }
        'Devices' {$AADRDevices = $true }
        'Default'
        {
            $AADRTenant = $true
            $AADRDomain = $true
            $AADRLicenses = $true
            $AADRUsers = $true
            $AADRServicePrincipals = $true
            $AADRDirectoryRoles = $true
            $AADRDirectoryRoleMembers = $true
            $AADRGroups = $true
            $AADRGroupMembers = $true
            $AADRDevices = $true

            If ($OutputType -eq "Default")
            {
                [array] $OutputType = "CSV","Excel"
            }
        }
    }

    Switch ($OutputType)
    {
        'STDOUT' { $AADRSTDOUT = $true }
        'CSV'
        {
            $AADRCSV = $true
            $AADRCreate = $true
        }
        'XML'
        {
            $AADRXML = $true
            $AADRCreate = $true
        }
        'JSON'
        {
            $AADRJSON = $true
            $AADRCreate = $true
        }
        'HTML'
        {
            $AADRHTML = $true
            $AADRCreate = $true
        }
        'Excel'
        {
            $AADRExcel = $true
            $AADRCreate = $true
        }
        'All'
        {
            #$ADRSTDOUT = $true
            $AADRCSV = $true
            $AADRXML = $true
            $AADRJSON = $true
            $AADRHTML = $true
            $AADRExcel = $true
            $AADRCreate = $true
            [array] $OutputType = "CSV","XML","JSON","HTML","Excel"
        }
        'Default'
        {
            [array] $OutputType = "STDOUT"
            $AADRSTDOUT = $true
        }
    }

    If ( ($AADRExcel) -and (-Not $AADRCSV) )
    {
        $AADRCSV = $true
        [array] $OutputType += "CSV"
    }

    $returndir = Get-Location
    $date = Get-Date

    # Create Output dir
    If ( ($AADROutputDir) -and ($AADRCreate) )
    {
        If (!(Test-Path $AADROutputDir))
        {
            New-Item $AADROutputDir -type directory | Out-Null
            If (!(Test-Path $AADROutputDir))
            {
                Write-Output "[Invoke-AzureADRecon] Error, invalid OutputDir Path ... Exiting"
                Return $null
            }
        }
        $AADROutputDir = $((Convert-Path $AADROutputDir).TrimEnd("\"))
        Write-Verbose "[*] Output Directory: $AADROutputDir"
    }
    ElseIf ($AADRCreate)
    {
        $AADROutputDir =  -join($returndir,'\','AzureADRecon-Report-',$(Get-Date -UFormat %Y%m%d%H%M%S))
        New-Item $AADROutputDir -type directory | Out-Null
        If (!(Test-Path $AADROutputDir))
        {
            Write-Output "[Invoke-AzureADRecon] Error, could not create output directory"
            Return $null
        }
        $AADROutputDir = $((Convert-Path $AADROutputDir).TrimEnd("\"))
        Remove-Variable AADRCreate
    }
    Else
    {
        $AADROutputDir = $returndir
    }

    If ($AADRCSV)
    {
        $CSVPath = [System.IO.DirectoryInfo] -join($AADROutputDir,'\','CSV-Files')
        New-Item $CSVPath -type directory | Out-Null
        If (!(Test-Path $CSVPath))
        {
            Write-Output "[Invoke-AzureADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable AADRCSV
    }

    If ($AADRXML)
    {
        $XMLPath = [System.IO.DirectoryInfo] -join($AADROutputDir,'\','XML-Files')
        New-Item $XMLPath -type directory | Out-Null
        If (!(Test-Path $XMLPath))
        {
            Write-Output "[Invoke-AzureADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable AADRXML
    }

    If ($AADRJSON)
    {
        $JSONPath = [System.IO.DirectoryInfo] -join($AADROutputDir,'\','JSON-Files')
        New-Item $JSONPath -type directory | Out-Null
        If (!(Test-Path $JSONPath))
        {
            Write-Output "[Invoke-AzureADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable AADRJSON
    }

    If ($AADRHTML)
    {
        $HTMLPath = [System.IO.DirectoryInfo] -join($AADROutputDir,'\','HTML-Files')
        New-Item $HTMLPath -type directory | Out-Null
        If (!(Test-Path $HTMLPath))
        {
            Write-Output "[Invoke-AzureADRecon] Error, could not create output directory"
            Return $null
        }
        Remove-Variable AADRHTML
    }

    # AzureAD Login

    If ($Method -eq 'AzureAD')
    {
        Try
        {
            Write-Output "[Invoke-AzureADRecon] AzureAD Module is installed. Logging in ..."
            If ($Credential -eq [Management.Automation.PSCredential]::Empty)
            {
                If ($TenantID)
                {
                    Connect-AzureAD  -TenantID $TenantID | Out-Null
                }
                Else
                {
                    Connect-AzureAD | Out-Null
                }
            }
            Else
            {
				 If ($TenantID)
                {
                    Connect-AzureAD -TenantID $TenantID -Credential $Credential | Out-Null
                }
                Else
                {
                    Connect-AzureAD -Credential $Credential | Out-Null
                }
            }
        }
        Catch
        {
            Write-Warning "[Invoke-AzureADRecon] Error authenticating to AzureAD ... Exiting"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            Return $null
        }
    }

    Write-Debug "AzureAD Logged In"

    Write-Output "[*] Commencing - $date"
    If ($AADRTenant)
    {
        Write-Output "[-] Tenant"
        $AADRObject = Get-AADRTenant -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "Tenant"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRTenant
    }
    If ($AADRDomain)
    {
        Write-Output "[-] Domain"
        $AADRObject = Get-AADRDomain -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "Domain"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRDomain
    }
    If ($AADRLicenses)
    {
        Write-Output "[-] Licenses"
        $AADRObject = Get-AADRLicense -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "Licenses"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRLicenses
    }
    If ($AADRUsers)
    {
        Write-Output "[-] Users - May take some time"
        $AADRObject = Get-AADRUser -Method $Method -date $date -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "Users"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRUsers
    }
    If ($AADRServicePrincipals)
    {
        Write-Output "[-] Service Principals - May take some time"
        $AADRObject = Get-AADRServicePrincipal -Method $Method -date $date -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "ServicePrincipals"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRServicePrincipals
    }
    If ($AADRDirectoryRoles)
    {
        Write-Output "[-] Directory Role"
        $AADRObject = Get-AADRDirectoryRole -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "DirectoryRoles"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRDirectoryRoles
    }
    If ($AADRDirectoryRoleMembers)
    {
        Write-Output "[-] Directory Role Membership - May take some time"
        $AADRObject = Get-AADRDirectoryRoleMember -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "DirectoryRoleMembers"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRDirectoryRoleMembers
    }
    If ($AADRGroups)
    {
        Write-Output "[-] Groups - May take some time"
        $AADRObject = Get-AADRGroup -Method $Method -date $date -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "Groups"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRGroups
    }
    If ($AADRGroupMembers)
    {
        Write-Output "[-] Group Membership - May take some time"
        $AADRObject = Get-AADRGroupMember -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "GroupMembers"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRGroupMembers
    }
    If ($AADRDevices)
    {
        Write-Output "[-] Devices - May take some time"
        $AADRObject = Get-AADRDevice -Method $Method -Threads $Threads
        If ($AADRObject)
        {
            Export-ADR -ADRObj $AADRObject -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "Devices"
            Remove-Variable AADRObject
        }
        Remove-Variable AADRDevices
    }

    $TotalTime = "{0:N2}" -f ((Get-DateDiff -Date1 (Get-Date) -Date2 $date).TotalMinutes)

    $AboutAzureADRecon = Get-AADRAbout -Method $Method -date $date -AzureADReconVersion $AzureADReconVersion -Credential $Credential -RanonComputer $RanonComputer -TotalTime $TotalTime

    If ( ($OutputType -Contains "CSV") -or ($OutputType -Contains "XML") -or ($OutputType -Contains "JSON") -or ($OutputType -Contains "HTML") )
    {
        If ($AboutAzureADRecon)
        {
            Export-ADR -ADRObj $AboutAzureADRecon -AADROutputDir $AADROutputDir -OutputType $OutputType -ADRModuleName "AboutAzureADRecon"
        }
        Write-Output "[*] Total Execution Time (mins): $($TotalTime)"
        Write-Output "[*] Output Directory: $AADROutputDir"
        $ADRSTDOUT = $false
    }

    Switch ($OutputType)
    {
        'STDOUT'
        {
            If ($ADRSTDOUT)
            {
                Write-Output "[*] Total Execution Time (mins): $($TotalTime)"
            }
        }
        'HTML'
        {
            Export-ADR -ADRObj $(New-Object PSObject) -AADROutputDir $AADROutputDir -OutputType $([array] "HTML") -ADRModuleName "Index"
        }
        'EXCEL'
        {
            Export-ADRExcel $AADROutputDir
        }
    }
    Remove-Variable TotalTime
    Remove-Variable AboutAzureADRecon
    Set-Location $returndir
    Remove-Variable returndir

    If ($Method -eq 'AzureAD')
    {
        Disconnect-AzureAD
    }

    If ($AADROutputDir)
    {
        Remove-EmptyAADROutputDir $AADROutputDir $OutputType
    }

    Remove-Variable AzureADReconVersion
    Remove-Variable RanonComputer
}

If ($Log)
{
    Start-Transcript -Path "$(Get-Location)\AzureADRecon-Console-Log.txt"
}

#$Credential = New-Object System.Management.Automation.PSCredential ("Username", $(ConvertTo-SecureString "Password" -AsPlainText -Force))

Invoke-AzureADRecon -GenExcel $GenExcel -Method $Method -Collect $Collect -Credential $Credential -OutputType $OutputType -AADROutputDir $OutputDir -Threads $Threads

If ($Log)
{
    Stop-Transcript
}
