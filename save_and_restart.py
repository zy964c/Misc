import time
import win32com.client
import os

def save_and_shutdown(file_path, file_name, pause):
    """
    Saves file after pause in seconds. 15 minutes given for CATIA to save.
    After that logs off the user
    """
    time.sleep(pause)
    catia = win32com.client.Dispatch('catia.application')
    documents = catia.Documents
    for doc in range(1, documents.Count + 1):
        if file_name in documents.Item(doc).Name and 'cgr' not in documents.Item(doc).Name:
            our_doc = documents.Item(doc)
            break
    our_doc.SaveAs(file_path + '\\' + file_name + '.CATPart')

    time.sleep(900)

    os.system("shutdown /l")
    
def log_user_off(pause):
    """
    Saves file after pause in seconds. 15 minutes given for CATIA to save.
    After that logs off the user
    """
    time.sleep(pause)
    
    os.system("shutdown /l")

def save():

    catia = win32com.client.Dispatch('catia.application')
    documents = catia.Documents
    for doc in range(1, documents.Count + 1):
        if 'CA' in documents.Item(doc).Name and 'cgr' not in documents.Item(doc).Name:
            doc1 = documents.Item(doc)
            doc1.Save
            #(r'C:\Temp\zy964c\new_carm.CATDocument')
            print 'saved'
if __name__ == "__main__":

    #save_and_shutdown('C:\Temp\zy964c', 'abcdefg', 60)
    save()
