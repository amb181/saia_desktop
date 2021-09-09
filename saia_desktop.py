from PyQt5 import QtCore, QtGui, QtWidgets, QtWebEngineWidgets
from zipfile import ZipFile
from subprocess import Popen

from office365.runtime.auth.client_credential import ClientCredential 
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

import sys, os, shutil, time, webbrowser, pyodbc

# # Check if latest version
def check_db():

    con = pyodbc.connect(   
                            """
                            DRIVER={ODBC Driver 17 for SQL Server};
                            SERVER=100.126.82.6;
                            DATABASE=DFA_EXPOSE;
                            PORT=1433;
                            UID=dfa_report_user;
                            PWD=DFA.rep0rt.308;
                            MARS_Connection=yes;
                            """
                        )
    cur = con.cursor()

    db_cmd =    """ 
                SELECT TOP (1) [version] 
                FROM [DFA_EXPOSE].[dbo].[exp_sql_saia_desktop_versions] 
                ORDER BY [version] DESC; 
                """
    results = cur.execute(db_cmd)

    for row in results:
        version = row[0]

    return version


# # Download file from sharepoint
def Download_zip():
    file_to_download = ""
    """ Connect to sharepoint """
    cc = ClientCredential("de0ff754-1ae1-436f-9b63-4ea4557a45a8", "0v4z5s80/KD34nB3sEZyPQAikjTXpO8mIKHQUzWvGUU=")
    ctx = ClientContext("https://ericssonnam.sharepoint.com/sites/SourcingAIAssistant").with_credentials(cc)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    path_list = 'Shared Documents/SourcingAIAssistant/'

    output_file_name = 'SAIA_desktop.zip'

    """ Get root folder from sharepoint """
    folder = ctx.web.lists.get_by_title('Documents').root_folder
    """ Creating query to obtain root folder with sharepoint connection """
    ctx.load(folder)
    """ Run query with sharepoint connection """
    ctx.execute_query()

    """ Navigate trough folders """
    for path in path_list:
        """ Activate sub folders from root """
        subfolders = folder.folders
        """ Query to load folders existing in directory using sharepoint connection """
        ctx.load(subfolders)
        """ Run query for sharepoint connection """
        ctx.execute_query()

        """ Search for subfolders """
        for subfolder in subfolders:
            """ Subfolder found """
            if(subfolder.properties['Name'] == path):
                """ Set subfolder as the new folder root """
                folder = subfolder

    """ Get files from subfolder """
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    """ Look for specific file """
    selectedFile = None
    """ This for loop is used to start counting months and iterate with get_file_date function """
    for file in files:
        """ File found """
        if(file.properties['Name'] == output_file_name):
            selectedFile = file
            file_to_download = selectedFile.properties['Name']
            print("File to download: ", file_to_download)
            break

    """ Use the original file name to notify that was not found """
    if selectedFile is None:
        raise ValueError("%s does not exist!" % output_file_name)

    """ Get full name of file location """
    completeName = os.path.join(os.environ['USERPROFILE'], 'Downloads', selectedFile.properties['Name'])
    """ Check if output_file_name has already downloaded """
    if os.path.exists(completeName):
        """ If it has been downloaded delete """
        os.remove(completeName)

    """ Download specific file """
    with open(completeName, 'wb') as output_file:
        response = File.open_binary(ctx, selectedFile.properties['ServerRelativeUrl'])
        output_file.write(response.content)

    print("File downloaded!")

    return file_to_download


# # Update window
def Update_window():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Information)
    msg.setWindowTitle("New version available!")
    msg.setText("The latest version will be installed.\nThis may take a few minutes.")
    msg.setWindowIcon(QtGui.QIcon('C:\\Program Files\\SAIA\\bin\\favicon.ico'))
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
    selection = msg.exec_()

    return selection

# # Class to open new links (override acceptNavigationRequest)
class WebEnginePage(QtWebEngineWidgets.QWebEnginePage):
    def acceptNavigationRequest(self, url,  _type, isMainFrame):
        if _type == QtWebEngineWidgets.QWebEnginePage.NavigationTypeLinkClicked:
            if "login.microsoftonline" not in str(url):
                link = str(url).replace("PyQt5.QtCore.QUrl('", "").replace("')", "")
                webbrowser.open(link)
            return True
        return super(WebEnginePage, self).acceptNavigationRequest(url, _type, isMainFrame)


# # Main class
class SAIA_chat(QtWebEngineWidgets.QWebEngineView):
    def __init__(self, windows, *args, **kwargs):
        super(SAIA_chat, self).__init__(*args, **kwargs)
        """ Window's properties """
        self.setWindowTitle('SAIA')
        self.setFixedSize(380, 545)
        self.setWindowIcon(QtGui.QIcon('C:\\Program Files\\SAIA\\bin\\favicon.ico'))
        """ Init window """
        self.setPage(WebEnginePage(self))
        self._windows = windows
        self._windows.append(self)
        self.page().settings().setAttribute(QtWebEngineWidgets.QWebEngineSettings.ShowScrollBars, False)        

    def closeEvent(self, event):
        """ if closed window's id is the same as the main window's """
        if str(self.window).split("0x", 1)[1] == str(self._windows[0]).split("0x",1)[1]:
            event.ignore()
            self.hide()
            """ Add system tray icon """
            self.trayIcon = QtWidgets.QSystemTrayIcon(QtGui.QIcon('C:\\Program Files\\SAIA\\bin\\favicon.ico'))
            self.trayIcon.setToolTip('Sourcing AI Assistant')
            self.trayIcon.show()
            self.trayIcon.showMessage("SAIA is still running!", "SAIA will continue running in the background" )
            """ Function for tray """
            self.menu = QtWidgets.QMenu()
            """ Re-open """
            self.reopen_action = self.menu.addAction('Open')        
            self.reopen_action.triggered.connect(self.Re_Open)
            """ Exit """
            self.exit_action = self.menu.addAction('Exit')          
            self.exit_action.triggered.connect(app.quit)  
            """ Add actions to menu """
            self.trayIcon.setContextMenu(self.menu)
            """ Action on-click """
            self.trayIcon.activated.connect(self.trayIconActivated) 

    """ Open with right click (menu) """
    def Re_Open(self):      
        self.show()
        self.raise_()
        self.trayIcon.hide()

    """ Open with left click (click on tray icon) """
    def trayIconActivated(self, reason):    
        if reason == self.trayIcon.Trigger:
            self.show()
            self.raise_()
            self.trayIcon.hide()

    """ Create and close new window when link is clicked, this will trigger the closeEvent, that's why there's an if there """
    def createWindow(self, _type):
        if QtWebEngineWidgets.QWebEnginePage.WebBrowserTab:
            v = SAIA_chat(self._windows)
            v.resize(640, 480)
            v.show()
            v.close()
            return v



# # #  MAIN FUNCTION  # # #
if __name__ == '__main__':
    local_v = '1.1'                                 # Local version
    prod_v = check_db()                             # DB latest version
    
    # Init QApplication
    app = QtWidgets.QApplication(sys.argv)
    
    if local_v == prod_v:                           # Local version equals db latest version
        # Delete downloaded zip and extracted folder
        path = 'C:\\Users\\' + os.getlogin() + '\\Downloads'
        if os.path.isfile(path + '\\SAIA_desktop.zip'):
            os.remove(path + '\\SAIA_desktop.zip')
        if os.path.isdir(path + '\\SAIA_desktop'):
            shutil.rmtree(path + '\\SAIA_desktop')

        windows = []
        w = SAIA_chat(windows)
        w.load(QtCore.QUrl("https://sourcingai.internal.ericsson.com/home.php"));
        w.show()
        sys.exit(app.exec_())
    
    else:
        # Show popup window
        sel = Update_window()
        if sel == 1024:                                                         # 'Ok' selected
            # Download zip file from sharepoint
            file_to_download = Download_zip()
            # Unzip file
            path = 'C:\\Users\\' + os.getlogin() + '\\Downloads'
            with ZipFile(path + '\\' + file_to_download, 'r') as zip_ref:
                zip_ref.extractall(path)
            # Execute batch file
            p = Popen('Setup_SAIA.bat', shell=True, cwd= path + '\\' + file_to_download.replace('.zip', ''))
            stdout, stderr = p.communicate()
        else:                                                               # 'Cancel' selected
            sys.exit()