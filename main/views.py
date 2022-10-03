import mimetypes
import os

from django.core.files.storage import FileSystemStorage
from django.http import HttpResponse
from django.shortcuts import render, redirect
from threading import Thread
from .yuval_shit import TimeSaver
import pyexcel

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# Create your views here.
def homepage(request):
    if request.method == 'POST':

        # data.xls
        myfile = request.FILES['myfile']
        fs = FileSystemStorage()
        filename = fs.save("data.xls", myfile)
        uploaded_file_url = BASE_DIR + fs.url(filename)

        pyexcel.save_book_as(file_name=uploaded_file_url,
                             dest_file_name=uploaded_file_url + "x")
        uploaded_file_url += "x"

        # others
        splited_names = {}
        current_leader = request.POST.get("leader1", False)
        index = 1
        while current_leader:
            splited_names[current_leader] = request.POST.get(f"children{index}", False).split("\r\n")
            index += 1
            current_leader = request.POST.get(f"leader{index}", False)

        # send it
        t = TimeSaver(uploaded_file_url, splited_names, request.POST.get(f"age", False))
        t.save_my_time()

        # remove data
        os.remove(uploaded_file_url)

        # download
        return download()

    return render(request=request,
                  template_name="main/home.html")


def download():
    filename = "result.xlsx"
    # Define the full file path
    filepath = BASE_DIR + "\\" + filename
    # Open the file for reading content
    path = open(filepath, 'rb')
    # Set the mime type
    mime_type, _ = mimetypes.guess_type(filepath)
    # Set the return value of the HttpResponse
    response = HttpResponse(path, content_type=mime_type)
    # Set the HTTP header for sending to browser
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    return response


def thanks(request, missed=""):
    print(missed)
    return render(request=request,
                  template_name="main/thanks.html")
