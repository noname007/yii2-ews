# yii2-ews
a simple wrapper for php-ew

### about service application account
ews docs: https://docs.microsoft.com/zh-cn/Exchange/client-developer/exchange-web-services/ews-application-types

### install
```shell
php composer.phar require noname007/yii2-ews
```

### config
```php
....

   component => [
        ...
        'ews' => [
            'class' => Ews::class,
            'host' => 'exchange serve domain',
            'password' => 'service application account',
            'username' => 'service application account',
        ]
   ]

...
..
```

### usage
```php
$ews = Yii::$app->ews;

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

```