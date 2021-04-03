<?php

namespace App\Http\Controllers;

use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Http\Request;
use Illuminate\Routing\Controller as BaseController;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Validator;
use Illuminate\Validation\ValidationException;
use PhpOffice\PhpWord\Element\Image;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\IOFactory;

class Controller extends BaseController
{
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    public function getWordPic(Request $request) {
        $params = $request->all();
        $rules = [
            'file' => 'bail|required|file',
        ];
        $messages = [
            'file.required' => '必须传输file参数',
            'file.file' => 'file参数必须为成功上传的文件',
        ];
        $validator = Validator::make($params, $rules, $messages);

        if ($validator->fails()) {
            $message = $validator->errors()->first();
            throw new ValidationException( $validator, response()->json([
                'code' => 1,
                'message' => $message,
            ], 422) );
        }

        //
        $date = date('Ymd');
        $file = $request->file('file');
        $sections = IOFactory::load($file)->getSections();

        foreach ($sections as $section) {
            $elements = $section->getElements();

            foreach ($elements as $element) {
                if ($element instanceof TextRun) {
                    $text_run_elements = $element->getElements();

                    foreach ($text_run_elements as $text_run_element) {
                        if ($text_run_element instanceof Image) {
                            $cache_key = 'new_pic_name' . $date . 'suffix';
                            $new_pic_name_suffix = Cache::has($cache_key) !== false ? Cache::get($cache_key) : 1;
                            $new_pic_name = 'auto_' . $date . '_' . $new_pic_name_suffix;
                            Cache::put($cache_key, $new_pic_name_suffix + 1, now()->addDay());

                            $image_data = $text_run_element->getImageStringData(true);
                            $image_src = storage_path('app/public/wordPic/'. $new_pic_name . '.' . $text_run_element->getImageExtension());
                            file_put_contents($image_src, base64_decode($image_data));
                        }
                    }

                }
            }

        }

        return response()->json([
            'code' => 0,
            'message' => '提取word文件中的图片成功',
        ]);
    }

}
