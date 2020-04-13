<?php


namespace App\Nova\Actions;

use Illuminate\Bus\Queueable;
use Laravel\Nova\Actions\Action;
use Illuminate\Support\Collection;
use Laravel\Nova\Fields\ActionFields;
use Illuminate\Queue\SerializesModels;
use Illuminate\Queue\InteractsWithQueue;
use Laravel\Nova\Fields\Text;
use Carbon\Carbon;

class DownloadReferenceWord extends Action
{
    use InteractsWithQueue, Queueable, SerializesModels;

    public $name = 'Справка  в формате  Word';


    /**
     * Perform the action on the given models.
     *
     * @param \Laravel\Nova\Fields\ActionFields $fields
     * @param \Illuminate\Support\Collection $models
     * @return mixed
     */
    public function handle(ActionFields $fields, Collection $models)
    {
// используем модель Position  -  и ее связи
        foreach ($models as $model) {
            $company_name = $model->company->name;// название компании
            $company_director = $model->company->director_fio_dp; //фамилия директора

            $company_details = $model->company->details; //реквизиты компании
            $user_famaly=$model->user->login;//фамилия выбранного юзера

            $user_name = $model->user->surname . " " . $model->user->name . " " . $model->user->patronymic; // ФИО выбранного
            $user_date = $model->user->birthday;//дата рождения
            $user_position = $model->position;//должность
        }
        // $now_date= date("m.d.y");
        $now_date= Carbon::now()->format('d.m.Y');
        $user_dat=Carbon::parse($user_date)->format('d.m.Y');

        // делаем прям тут  создание ворд документа по  шаблону с последующей выгрузкой
        $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor(storage_path('app/doc_templates/shablon_shablonspravka.docx'));
        $templateProcessor->setValue('company_name',  $company_name);
        $templateProcessor->setValue('company_director', $company_director);
        $templateProcessor->setValue('company_details', $company_details );
        $templateProcessor->setValue('user_name',$user_name);
        $templateProcessor->setValue('user_date',  $user_dat);
        $templateProcessor->setValue('user_position', $user_position);
        $templateProcessor->setValue('now_date', $now_date);
        $output = storage_path('app/public/'.'shablon_shablonspravka' . '_' .$user_famaly. '.docx'); //  сохраняем документпо фамилии юзера

        $templateProcessor->saveAs($output);

        return Action::download(
            url('/reference_onepeople') . '/?' . http_build_query([
                'path'     => storage_path("app/public/"),
                'filename' => 'shablon_shablonspravka' . '_' .$user_famaly. '.docx',

            ]),
            'farpost_shablonspravka' . '_' .$user_famaly. '.docx'
        );
    }

    /**
     * Get the fields available on the action.
     *
     * @return array
     */
    public function fields()
    {
        return [];
    }


}

