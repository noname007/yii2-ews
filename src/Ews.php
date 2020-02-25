<?php


namespace noname007\yii2ews;

use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAttendeesType;
use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use jamesiarmes\PhpEws\Enumeration\CalendarItemCreateOrDeleteOperationType;
use jamesiarmes\PhpEws\Enumeration\ResponseClassType;
use jamesiarmes\PhpEws\Enumeration\RoutingType;
use jamesiarmes\PhpEws\Request\CreateItemType;
use jamesiarmes\PhpEws\Type\AttendeeType;
use jamesiarmes\PhpEws\Type\BodyType;
use jamesiarmes\PhpEws\Type\CalendarItemType;
use jamesiarmes\PhpEws\Type\ConnectingSIDType;
use jamesiarmes\PhpEws\Type\EmailAddressType;
use jamesiarmes\PhpEws\Type\ExchangeImpersonationType;
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
     * @var Client
     */
    private $client;

    /**
     * @return Client
     */
    public function getClient()
    {
        if(!$this->client)
        {
            $client = new \jamesiarmes\PhpEws\Client($this->host, $this->username, $this->password, $this->version);

            $client->setTimezone($this->timezone);

            $this->setClient($client);
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
    public function createAppointment(\DateTime $start, \DateTime $end, $subject, $guests = [], $body = '', $body_type = BodyTypeType::TEXT)
    {
        //Build the request,
        $request = new CreateItemType();
        $request->SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType::SEND_ONLY_TO_ALL;
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
                \Yii::error("Event failed to create with \"$code: $message\"\n");
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
}
