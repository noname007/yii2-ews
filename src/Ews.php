<?php


namespace noname007\yii2ews;

use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAttendeesType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfBaseItemIdsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfItemChangeDescriptionsType;
use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use jamesiarmes\PhpEws\Enumeration\CalendarItemCreateOrDeleteOperationType;
use jamesiarmes\PhpEws\Enumeration\CalendarItemUpdateOperationType;
use jamesiarmes\PhpEws\Enumeration\ConflictResolutionType;
use jamesiarmes\PhpEws\Enumeration\DisposalType;
use jamesiarmes\PhpEws\Enumeration\ResponseClassType;
use jamesiarmes\PhpEws\Enumeration\RoutingType;
use jamesiarmes\PhpEws\Enumeration\UnindexedFieldURIType;
use jamesiarmes\PhpEws\Request\CreateItemType;
use jamesiarmes\PhpEws\Request\DeleteItemType;
use jamesiarmes\PhpEws\Request\ExpandDLType;
use jamesiarmes\PhpEws\Request\ResolveNamesType;
use jamesiarmes\PhpEws\Request\UpdateItemType;
use jamesiarmes\PhpEws\Type\AttendeeType;
use jamesiarmes\PhpEws\Type\BodyType;
use jamesiarmes\PhpEws\Type\CalendarItemType;
use jamesiarmes\PhpEws\Type\ConnectingSIDType;
use jamesiarmes\PhpEws\Type\EmailAddressType;
use jamesiarmes\PhpEws\Type\ExchangeImpersonationType;
use jamesiarmes\PhpEws\Type\ItemChangeType;
use jamesiarmes\PhpEws\Type\ItemIdType;
use jamesiarmes\PhpEws\Type\PathToUnindexedFieldType;
use jamesiarmes\PhpEws\Type\SetItemFieldType;
use noname007\yii2ews\models\Guests;
use yii\base\Component;

class Ews extends Component
{
    protected $host;

    protected $username;

    protected $password;

    protected $timezone = 'China Standard Time';

    protected $version = Client::VERSION_2016;

    /**
     * too long  result to a 940 seconds request on the next ews op request,
     * need to be renewed
     *
     * @var int client longest live time
     */
    protected $client_llt = 600;

    /**
     * @param int $client_llt
     */
    public function setClientLlt($client_llt)
    {
        $this->client_llt = $client_llt;
    }


    /**
     * @var int expire at the {$expiration} (timestamps)
     */
    private $expiration = 0;

    /**
     * @var Client
     */
    private $client;

    protected $curl_opt = [];

    /**
     * @param array $curl_opt
     */
    public function setCurlOpt($curl_opt)
    {
        $this->curl_opt = $curl_opt;
    }

    
    /**
     * @return Client
     */
    public function getClient()
    {
        $now = time();
        if($this->expiration < $now || !$this->client)
        {
            $client = new \jamesiarmes\PhpEws\Client($this->host, $this->username, $this->password, $this->version);
            $client->setCurlOptions($this->curl_opt);
            $client->setTimezone($this->timezone);

            $this->setClient($client);
            $this->expiration = $now + $this->client_llt;
        }

        return $this->client;
    }

    /**
     * @param mixed $host
     */
    public function setHost($host)
    {
        $this->host = $host;
    }

    /**
     * @param mixed $user
     */
    public function setUsername($username)
    {
        $this->username = $username;
    }

    /**
     * @param mixed $password
     */
    public function setPassword($password)
    {
        $this->password = $password;
    }

    /**
     * @param string $timezone
     */
    public function setTimezone($timezone)
    {
        $this->timezone = $timezone;
    }

    /**
     * @param string $version
     */
    public function setVersion($version)
    {
        $this->version = $version;
    }

    /**
     * @param Client $client
     */
    public function setClient($client)
    {
        $this->client = $client;
    }

    public function impersonateByPrimarySmtpAddress($email)
    {
        $sid = new ConnectingSIDType();
        $sid->PrimarySmtpAddress = $email;
        $impersonation = new ExchangeImpersonationType();
        $impersonation->ConnectingSID = $sid;
        $this->getClient()->setImpersonation($impersonation);
    }

    /**
     * @docs https://docs.microsoft.com/zh-cn/Exchange/client-developer/web-service-reference/createitem-operation-calendar-item
     *
     * @param \DateTime $start
     * @param \DateTime $end
     * @param           $subject
     * @param bool     $is_meeting
     * @param Guests[] $guests
     * @param string    $body
     * @param string    $body_type
     *
     * @return string|null
     */
    public function createAppointment(\DateTime $start, \DateTime $end, $subject, $guests = [], $body = '', $body_type = BodyTypeType::TEXT, callable  $moidify_request_call = null)
    {
        //Build the request,
        $request = new CreateItemType();

        $request->SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType::SEND_TO_ALL_AND_SAVE_COPY;
        $request->Items = new NonEmptyArrayOfAllItemsType();

        // Build the event to be added.
        $event = new CalendarItemType();
        $event->RequiredAttendees = new NonEmptyArrayOfAttendeesType();

        //$event->IsMeeting = $is_meeting;
        //error

        $event->Start = $start->format('c');
        $event->End = $end->format('c');
        $event->Subject = $subject;

        // Set the event body.
        $event->Body = new BodyType();
        $event->Body->_ = $body;
        $event->Body->BodyType = $body_type;


        // Iterate over the guests, adding each as an attendee to the request.
        foreach ($guests as $guest) {
            $attendee = new AttendeeType();
            $attendee->Mailbox = new EmailAddressType();
            $attendee->Mailbox->EmailAddress = $guest['email'];
            $attendee->Mailbox->Name = $guest['name'];
            $attendee->Mailbox->RoutingType = RoutingType::SMTP;
            $event->RequiredAttendees->Attendee[] = $attendee;
        }

        if(is_callable($moidify_request_call)) {
            if(!call_user_func($moidify_request_call, $request, $event)) {
                return null;
            }
        }

        // Add the event to the request. You could add multiple events to create more
        // than one in a single request.
        $request->Items->CalendarItem[] = $event;

        $response = $this->getClient()->CreateItem($request);


        // Iterate over the results, printing any error messages or event ids.
        $response_messages = $response->ResponseMessages->CreateItemResponseMessage;
        $id = null;
        foreach ($response_messages as $response_message) {
            // Make sure the request succeeded.
            if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
                $code = $response_message->ResponseCode;
                $message = $response_message->MessageText;
                \Yii::error([
                    'ResponseCode' => $code,
                    'MessageText' => $message,
                ]);
                continue;
            }

            // Iterate over the created events, printing the id for each.
            foreach ($response_message->Items->CalendarItem as $item) {
                $id = $item->ItemId->Id;
                \Yii::info("Created event $id");
            }
        }

        return $id;
    }

    /**
     * @docs https://docs.microsoft.com/zh-cn/Exchange/client-developer/web-service-reference/updateitem-operation
     * @docs https://docs.microsoft.com/zh-cn/Exchange/client-developer/web-service-reference/setitemfield
     * @param               $item_id
     * @param \DateTime     $start
     * @param \DateTime     $end
     * @param               $subject
     * @param array         $guests
     * @param string        $body
     * @param string        $body_type
     * @param callable|null $moidify_request_call
     *
     * @return bool|string|null
     */
    public function updateAppointment($item_id, \DateTime $start, \DateTime $end, $subject, $guests = [], $body = '', $body_type = BodyTypeType::TEXT, callable  $moidify_request_call = null) {

        $request = new UpdateItemType();
        $request->ConflictResolution = ConflictResolutionType::ALWAYS_OVERWRITE;
        $request->SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType::SEND_TO_ALL_AND_SAVE_COPY;


            $change = new ItemChangeType();
            $change->ItemId = new ItemIdType();
            $change->ItemId->Id = $item_id;
            $change->Updates = new NonEmptyArrayOfItemChangeDescriptionsType();

            // Set the updated start time.
            $field = new SetItemFieldType();
            $field->FieldURI = new PathToUnindexedFieldType();
            $field->FieldURI->FieldURI = UnindexedFieldURIType::CALENDAR_START;
            $field->CalendarItem = new CalendarItemType();
            $field->CalendarItem->Start = $start->format('c');
            $change->Updates->SetItemField[] = $field;

            // Set the updated end time.
            $field = new SetItemFieldType();
            $field->FieldURI = new PathToUnindexedFieldType();
            $field->FieldURI->FieldURI = UnindexedFieldURIType::CALENDAR_END;
            $field->CalendarItem = new CalendarItemType();
            $field->CalendarItem->End = $end->format('c');
            $change->Updates->SetItemField[] = $field;


            $field = new SetItemFieldType();
            $field->FieldURI = new PathToUnindexedFieldType();
            $field->FieldURI->FieldURI = UnindexedFieldURIType::ITEM_SUBJECT;
            $field->CalendarItem = new CalendarItemType();
            $field->CalendarItem->Subject = $subject;

            $change->Updates->SetItemField[] = $field;

            $field = new SetItemFieldType();
            $field->FieldURI = new PathToUnindexedFieldType();
            $field->FieldURI->FieldURI = UnindexedFieldURIType::ITEM_BODY;
            $field->CalendarItem = new CalendarItemType();
            $field->CalendarItem->Body = new BodyType();
            $field->CalendarItem->Body->_ = $body;
            $field->CalendarItem->Body->BodyType = $body_type;

            $change->Updates->SetItemField[] = $field;


            $field = new SetItemFieldType();
            $field->FieldURI = new PathToUnindexedFieldType();
            $field->FieldURI->FieldURI = UnindexedFieldURIType::CALENDAR_REQUIRED_ATTENDEES;
            $field->CalendarItem = new CalendarItemType();
            $field->CalendarItem->RequiredAttendees = new NonEmptyArrayOfAttendeesType();

            foreach ($guests as $guest) {
                $attendee = new AttendeeType();
                $attendee->Mailbox = new EmailAddressType();
                $attendee->Mailbox->EmailAddress = $guest['email'];
                $attendee->Mailbox->Name = $guest['name'];
                $attendee->Mailbox->RoutingType = RoutingType::SMTP;
                $field->CalendarItem->RequiredAttendees->Attendee[] = $attendee;
            }

            $change->Updates->SetItemField[] = $field;

            if($moidify_request_call) {
                if(!call_user_func($moidify_request_call, $request, $change)){
                    return false;
                }
            }

        $request->ItemChanges[] = $change;

        $response = $this->getClient()->UpdateItem($request);

        $id = null;
        $response_messages = $response->ResponseMessages->UpdateItemResponseMessage;

        foreach ($response_messages as $response_message) {
            // Make sure the request succeeded.
            if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
                $code = $response_message->ResponseCode;
                $message = $response_message->MessageText;
                \Yii::error([
                    'ResponseCode' => $code,
                    'MessageText' => $message,
                ]);
                continue;
            }

            // Iterate over the updated events, printing the id of each.
            foreach ($response_message->Items->CalendarItem as $item) {
                $id = $item->ItemId->Id;
                \Yii::info("Updated event $id");
            }
        }

        return $id;
    }

    /**
     *
     * @docs https://docs.microsoft.com/zh-cn/Exchange/client-developer/web-service-reference/deleteitem-operation
     *
     * @param $item_id
     *
     * @return bool
     */
    protected function deleteItem($item_id)
    {
        return $this->macro_delete_item($item_id);
    }

    public function cancelAppointment($item_id)
    {
        return
            $this->macro_delete_item(
                $item_id,
                function (DeleteItemType $deleteRequest) {
                    $deleteRequest->SendMeetingCancellations
                        = CalendarItemCreateOrDeleteOperationType::SEND_TO_ALL_AND_SAVE_COPY;
                })
            ;
    }

    /**
     * @docs https://docs.microsoft.com/zh-cn/exchange/client-developer/web-service-reference/resolvenames-operation?redirectedfrom=MSDN
     * @param               $entry
     * @param callable|null $moidify_request_call
     *
     * @return []
     */
    public function resolveNames($entry, callable  $moidify_request_call = null) {

        $client = $this->getClient();
        $request = new ResolveNamesType();

        $request->UnresolvedEntry = $entry;

        if(is_callable($moidify_request_call)) {
            if(!call_user_func($moidify_request_call, $request)) {
                return null;
            }
        }

        $response = $client->ResolveNames($request);

        $response_messages = $response->ResponseMessages->ResolveNamesResponseMessage;

        $result = [];

        foreach ($response_messages as $response_message) {
            // Make sure the request succeeded.
            if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
                $code = $response_message->ResponseCode;
                $message = $response_message->MessageText;

                \Yii::error([
                    'ResponseCode' => $code,
                    'MessageText' => $message,
                ]);
                continue;
            }

            // Iterate over the created events, printing the id for each.
            foreach ($response_message->ResolutionSet->Resolution as $resolutionType) {
                if($resolutionType->Mailbox) {
                    $result[] = [
                        'name' => $resolutionType->Mailbox->Name,
                        'email' => $resolutionType->Mailbox->EmailAddress,
                    ];
                }else if($resolutionType->Contact) {
                    //not found this type result
                }
            }
        }

        return $result;
    }

    /**
     * @docs https://docs.microsoft.com/zh-cn/exchange/client-developer/web-service-reference/expanddl-operation?redirectedfrom=MSDN
     * @param               $account
     * @param callable|null $moidify_request_call
     *
     * @return array|null
     */
    public function expandDl($account, callable  $moidify_request_call = null)
    {
        $client = $this->getClient();

        $request = new ExpandDLType();

        $request->Mailbox = new EmailAddressType();
        $request->Mailbox->EmailAddress = $account;

        if(is_callable($moidify_request_call)) {
            if(!call_user_func($moidify_request_call, $request)) {
                return null;
            }
        }

        $response =  $client->ExpandDL($request);


        $response_messages = $response->ResponseMessages->ExpandDLResponseMessage;

        $result = [];
        foreach ($response_messages as $response_message) {
            // Make sure the request succeeded.
            if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
                $code = $response_message->ResponseCode;
                $message = $response_message->MessageText;

                \Yii::error([
                    'ResponseCode' => $code,
                    'MessageText' => $message,
                ]);
                continue;
            }
            // Iterate over the created events, printing the id for each.
            foreach ($response_message->DLExpansion->Mailbox as $resolutionType) {
                $result[] = [
                    'name' => $resolutionType->Name,
                    'email' => $resolutionType->EmailAddress,
                    'mailbox_type' => $resolutionType->MailboxType,
                ];
            }
        }

        return $result;

    }

    protected function macro_delete_item($item_id, callable $callfunc =null)
    {
        $deleteRequest = new DeleteItemType();
        $deleteRequest->DeleteType = DisposalType::MOVE_TO_DELETED_ITEMS;

        if(is_callable($callfunc))
        {
            call_user_func($callfunc, $deleteRequest);
        }

        $deleteRequest->ItemIds
            = new NonEmptyArrayOfBaseItemIdsType();

        $item = new ItemIdType();
        $item->Id = $item_id;
        $deleteRequest->ItemIds->ItemId[] = $item;

        $response = $this->getClient()->DeleteItem($deleteRequest);

        $response_messages = $response->ResponseMessages->DeleteItemResponseMessage;

        foreach ($response_messages as $responseMessage)
        {
            if($responseMessage->ResponseClass == ResponseClassType::SUCCESS){
                return true;
            }else{
                $err_msg = [
                    'ResponseCode' => $responseMessage->ResponseCode,
                    'DescriptiveLinkKey' => $responseMessage->DescriptiveLinkKey,
                    'MessageText' => $responseMessage->MessageText,
                ];
                \Yii::error($err_msg);
                return false;
            }
        }
        return null;
    }
}
