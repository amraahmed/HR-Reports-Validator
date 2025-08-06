import os
from django.shortcuts import render
from django.http import FileResponse, JsonResponse
from django.conf import settings
from .excel_validator import validate_excel_file

def upload_file(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']

        # Save the uploaded file
        upload_path = os.path.join(settings.MEDIA_ROOT, excel_file.name)
        with open(upload_path, 'wb+') as destination:
            for chunk in excel_file.chunks():
                destination.write(chunk)

        # Generate validation report
        output_filename = f"validation_report_{os.path.splitext(excel_file.name)[0]}.txt"
        output_path = os.path.join(settings.MEDIA_ROOT, output_filename)

        try:
            validate_excel_file(upload_path, output_path)
            if os.path.exists(output_path):
                response = FileResponse(open(output_path, 'rb'))
                response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
                return response
            else:
                return JsonResponse({'error': 'Report file not generated'}, status=500)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=400)

    # For GET requests, render the upload form
    return render(request, 'validator/upload.html')

def download_file(request, filename):
    file_path = os.path.join(settings.MEDIA_ROOT, filename)
    if os.path.exists(file_path):
        response = FileResponse(open(file_path, 'rb'))
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    return HttpResponse('File not found', status=404) 