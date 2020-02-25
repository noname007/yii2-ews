<?php

use noname007\yii2ews\Ews;
use noname007\yii2ews\models\Guests;

require_once __DIR__.'/../vendor/autoload.php';
require(__DIR__ . '/../vendor/yiisoft/yii2/Yii.php');


Yii::setAlias('@noname007/yii2ews', __DIR__.'/../src');
/**
 * @var  $ews Ews
 */
$ews = Yii::createObject([
    'class' => Ews::class,
    'host' => 'exchange serve domain',
    'password' => 'service application account',
    'username' => 'service application account',
]);

$ews->impersonateByPrimarySmtpAddress('impersonated people email');

$guests =[
    new Guests(
        array('name' => 'John', 'email' => 'noname007@githubc.com',)
    ),
];

$ews->createAppointment(new DateTime("@".(time() + 15 * 60)),
    new DateTime('@'.(time() + 30 * 60)),
    'subject text',
    $guests
);