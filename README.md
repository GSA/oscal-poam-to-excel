# OSCAL-Conversion for Plan of Action and Milestones
This application renders an OSCAL XML POAM (Plan of Action Miestones) file to a Microsoft Excel file.  Currently the application can support a valid OSCAL Release 1.0.0 OSCAL XML document for the POAM.   The application does not reference other elements that the POAM is dependent upon that exist in other layers of the OSCAL component model (Implementation layer or Controls Layer).  It only populates elements that can be mapped directly from the Assessment layer.
# Why the project is useful
NIST is developing the Open Security Controls Assessment Language (OSCAL), a set of hierarchical, formatted, XML- and JSON-based formats that provide a standardized representation for different categories of information pertaining to the publication, implementation, and assessment of security controls. OSCAL is being developed through a collaborative approach with the public. The OSCAL website (https://csrc.nist.gov/Projects/Open-Security-Controls-Assessment-Language) provides an overview of the OSCAL project, including an XML and JSON schema reference and examples.
This minimal viable product application automates the manual mappings from OSCAL POAM to MS Excel.  This application product is meant only to provide an example on how OSCAL can be converted to the target document format and is not meant to be a production application.
# Project Requirements
The application is coded in C# as an ASP.NET web application Visual Studio project and is meant to run in standalone mode ONLY.   This project utilizes the OpenXML (DocumentFormat.OpenXml) and the Microsoft Office Interop (Microsoft.Office.Interop.Word) namespaces to perform XML parsing and perform document rendering.
# System Software Requirements
Windows 10
Office 2016 or better
Visual Studio 2019 Community Edition    
SQL Server Express 2019 or newer

# Getting started with this project
1. Install Visual Studio 2019 Community Edition
2. Clone and checkout the Project("https://github.com/GSA/oscal-poam-to-word”)
3. Clean and Build the Project
4. Locate and decompress the SQL back up file called “CONVERTREPO.zip” located under the DB folder in the project.
5. Open SQL Management Studio and restore the CONVERTREPO.bak to your local workstation SQL Server Express instance.
6.  Modify web.config file in main project folder to connect to your instance of SQL Express and the CONVERTREPO DB.
7. Run the Project.
8. Select a valid OSCAL XML POAM file and click "Upload"
9. The document rendering may take several minutes to complete depending upon the quantity of data contained in the originating XML document and will be made available for download after rendering is complete.   

# Known Issues
You may have the following error when running the code for the first time.
"Could not find a part of the path '\OSCAL-Conversion\OSCAL POAM Converter\bin\roslyn\csc.exe'".
This issue can be resolved by installing the Nuget Package with Visual Studio:  Microsoft.CodeDom.Providers.DotNetComplierPlatform.NoInstall
# License Information
This project is being released under the GNU General Public License (GNU GPL) 3.0 model. Under that model there is no warranty on the open source code.   Software under the GPL may be run for all purposes, including commercial purposes and even as a tool for creating proprietary software.
# Disclaimer
This project will ONLY work with Release 1.0.0 and some of the Milestone releases of the OSCAL schema from the NIST website (https://csrc.nist.gov/Projects/Open-Security-Controls-Assessment-Language) and does not address gaps between the NIST schema and the FedRAMP template requirements.   It maps elements between the two that match and will ignore those that do not.  
# Getting help with this project
Contact the GSA FedRAMP Project Management Office for more information or support.
# Originator of Code
VITG, INC.  http://www.volpegroup.com

