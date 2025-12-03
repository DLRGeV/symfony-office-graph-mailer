<?php

namespace Vitrus\SymfonyOfficeGraphMailer\Transport;

use Psr\EventDispatcher\EventDispatcherInterface;
use Psr\Log\LoggerInterface;
use Symfony\Component\Mailer\Envelope;
use Symfony\Component\Mailer\Exception\HttpTransportException;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mailer\Transport\AbstractApiTransport;
use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\Header\Headers;
use Symfony\Component\Mime\Header\ParameterizedHeader;
use Symfony\Component\Mime\Part\DataPart;
use Symfony\Contracts\HttpClient\Exception\TransportExceptionInterface;
use Symfony\Contracts\HttpClient\HttpClientInterface;
use Symfony\Contracts\HttpClient\ResponseInterface;

/**
 * @author Sjoerd Adema <vitrus@gmail.com>
 */
class GraphApiTransport extends AbstractApiTransport
{
    private static $accessTokenCache = [];

    private string $graphTenantId;
    private string $graphClientId;
    private string $graphClientSecret;

    public function __construct(
        string $graphTenantId,
        string $graphClientId,
        string $graphClientSecret,
        ?HttpClientInterface $client = null,
        ?EventDispatcherInterface $dispatcher = null,
        ?LoggerInterface $logger = null
    ) {
        $this->graphTenantId = $graphTenantId;
        $this->graphClientId = $graphClientId;
        $this->graphClientSecret = $graphClientSecret;

        parent::__construct($client, $dispatcher, $logger);
    }

    public function __toString(): string
    {
        return sprintf('microsoft-graph-api://%s:{SECRET}@%s', $this->graphClientId, $this->graphTenantId);
    }

    protected function doSendApi(SentMessage $sentMessage, Email $email, Envelope $envelope): ResponseInterface
    {
        $bodyStream = $this->convertToBase64Stream($email);

        $response = $this->client->request('POST', $this->getEndpoint($sentMessage), [
            'body' => $bodyStream,
            'auth_bearer' => $this->getAccessToken(),
        ]);

        try {
            $statusCode = $response->getStatusCode();
        } catch (TransportExceptionInterface $e) {
            throw new HttpTransportException('Could not reach Microsoft Graph API.', $response, 0, $e);
        } finally {
            fclose($bodyStream);
        }

        if (202 !== $statusCode) {
            throw new HttpTransportException('Unable to sent e-mail using Graph API', $response);
        }

        return $response;
    }

    /**
     * @return resource
     */
    private function convertToBase64Stream(Email $email)
    {
        $stream = fopen('php://temp', 'r+b');
        stream_filter_append($stream, 'convert.base64-encode', STREAM_FILTER_WRITE);

        foreach ($email->toIterable() as $chunk) {
            fwrite($stream, $chunk);
        }

        fflush($stream);
        rewind($stream);

        return $stream;
    }

    private function getAccessToken(): string
    {
        $cacheKey = "{$this->graphTenantId}:{$this->graphClientId}";
        $accessToken = self::$accessTokenCache[$cacheKey] ?? null;

        if($accessToken === null) {
            $accessToken = $this->requestAccessToken();
            self::$accessTokenCache[$cacheKey] = $accessToken;
        }

        return $accessToken;
    }

    private function requestAccessToken(): string
    {
        $url = 'https://login.microsoftonline.com/' . $this->graphTenantId . '/oauth2/v2.0/token';

        $response = $this->client->request('POST', $url, [
            'body' => [
                'client_id' => $this->graphClientId,
                'client_secret' => $this->graphClientSecret,
                'scope' => 'https://graph.microsoft.com/.default',
                'grant_type' => 'client_credentials',
            ],
        ]);

        $token = json_decode($response->getContent(), null, 512, JSON_THROW_ON_ERROR);
        return $token->access_token;
    }

    private function getEndpoint(SentMessage $sentMessage): string
    {
        $senderAddress = $sentMessage->getEnvelope()->getSender()->getAddress();

        return sprintf('https://graph.microsoft.com/v1.0/users/%s/sendMail', $senderAddress);
    }
}
