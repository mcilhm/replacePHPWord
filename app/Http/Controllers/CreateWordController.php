<?php

namespace App\Http\Controllers;

use App\ActivityTemplate;

class CreateWordController extends Controller
{
    //

    public function index()
    {

        $activity = ActivityTemplate::latest()->first();
        $phpWord = new \PhpOffice\PhpWord\PhpWord();

        $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor($activity->file);
        $values = array(
            array(
                'noreg'        => 1,
                'employee_name' => 'James',
                'dept_name'      => 'Taylor',
                'end_date'     => '+1 428 889 773',
            ),
            array(
                'noreg'        => 2,
                'employee_name' => 'Robert',
                'dept_name'      => 'Bell',
                'end_date'     => '+1 428 889 774',
            ),
            array(
                'noreg'        => 3,
                'employee_name' => 'Michael',
                'dept_name'      => 'Ray',
                'end_date'     => '+1 428 889 775',
            ),
        );
        $templateProcessor->cloneRowAndSetValues('noreg', $values);

        // dd($template);
        $pathFile = 'upload/file/';
        $fileName = time() . '.docx';

        $templateProcessor->saveAs($pathFile . $fileName);


        $headers = ['Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
        $newName = 'nicesnippets-pdf-file-' . time() . '.docx';

        return response()->download(public_path($pathFile . $fileName), $newName, $headers);

        // return Storage::download(public_path($pathFile . $fileName));
        // redirect(public_path($pathFile . $fileName));

        // response()->download(public_path($pathFile . $fileName), $fileName, ['Content-Type' => 'application/docx']);
    }
}
