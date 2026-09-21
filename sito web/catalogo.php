<?php

declare(strict_types=1);

/*
 * ============================================================================
 * CONFIGURAZIONE
 * ============================================================================
 */

$allowedHost = 'www.cartelli.it';
$defaultUrl  = 'https://www.cartelli.it/';


/*
 * ============================================================================
 * FUNZIONI URL
 * ============================================================================
 */

/**
 * Trasforma un URL relativo in un URL assoluto.
 */
function absoluteUrl(string $url, string $base): string
{
    $url = trim($url);

    if ($url === '') {
        return '';
    }

    /*
     * Anchor.
     */
    if (str_starts_with($url, '#')) {
        return $url;
    }

    /*
     * Schemi che non devono essere proxati.
     */
    if (preg_match(
        '~^(data|blob|javascript|mailto|tel):~i',
        $url
    )) {
        return $url;
    }

    /*
     * URL protocol-relative:
     *
     * //www.example.com/file.css
     */
    if (str_starts_with($url, '//')) {

        $baseParts = parse_url($base);

        return ($baseParts['scheme'] ?? 'https') . ':' . $url;
    }

    /*
     * URL assoluto.
     */
    if (preg_match('~^https?://~i', $url)) {
        return $url;
    }

    $baseParts = parse_url($base);

    $scheme = $baseParts['scheme'] ?? 'https';
    $host   = $baseParts['host'] ?? '';

    $basePath = $baseParts['path'] ?? '/';

    /*
     * URL dalla root:
     *
     * /images/logo.png
     */
    if (str_starts_with($url, '/')) {

        return $scheme . '://' .
               $host .
               $url;
    }

    /*
     * URL relativo:
     *
     * images/logo.png
     * ./images/logo.png
     * ../images/logo.png
     */

    $baseDir = dirname($basePath);

    $combined = $baseDir . '/' . $url;

    $segments = explode('/', $combined);

    $normalized = [];

    foreach ($segments as $segment) {

        if ($segment === '' || $segment === '.') {
            continue;
        }

        if ($segment === '..') {

            if (!empty($normalized)) {
                array_pop($normalized);
            }

            continue;
        }

        $normalized[] = $segment;
    }

    return $scheme . '://' .
           $host .
           '/' .
           implode('/', $normalized);
}


/**
 * Verifica che l'URL appartenga al dominio autorizzato.
 */
function isAllowedUrl(
    string $url,
    string $allowedHost
): bool {

    $parts = parse_url($url);

    if (!$parts || empty($parts['host'])) {
        return false;
    }

    $scheme = strtolower(
        $parts['scheme'] ?? ''
    );

    $host = strtolower(
        $parts['host']
    );

    return in_array(
        $scheme,
        ['http', 'https'],
        true
    ) && $host === strtolower($allowedHost);
}


/**
 * Genera l'URL locale del proxy.
 */
function proxyUrl(string $url): string
{
    return 'catalogo.php?url=' .
           rawurlencode($url);
}


/**
 * Trasforma un URL relativo nell'URL del proxy.
 */
function rewriteUrl(
    string $url,
    string $baseUrl,
    string $allowedHost
): string {

    $url = trim($url);

    /*
     * URL che non devono essere modificati.
     */
    if (
        $url === '' ||
        str_starts_with($url, '#') ||
        preg_match(
            '~^(data|blob|javascript|mailto|tel):~i',
            $url
        )
    ) {
        return $url;
    }

    $absolute = absoluteUrl(
        $url,
        $baseUrl
    );

    /*
     * Proxy solamente il dominio autorizzato.
     */
    if (!isAllowedUrl(
        $absolute,
        $allowedHost
    )) {
        return $url;
    }

    return proxyUrl($absolute);
}


/*
 * ============================================================================
 * COOKIE
 * ============================================================================
 */

/**
 * Converte l'header Set-Cookie remoto in cookie da restituire al browser.
 *
 * Vengono rimossi Domain, Path, Secure, HttpOnly e SameSite del server
 * remoto perché il cookie deve appartenere al proxy.
 */
function rewriteSetCookieForProxy(
    string $setCookie
): string {

    $parts = preg_split(
        '/;\s*/',
        $setCookie
    );

    if (!$parts || empty($parts[0])) {
        return $setCookie;
    }

    $result = [
        $parts[0]
    ];

    foreach (
        array_slice($parts, 1)
        as $part
    ) {

        $lower = strtolower(
            trim($part)
        );

        /*
         * Il dominio remoto non deve essere mantenuto.
         */
        if (str_starts_with(
            $lower,
            'domain='
        )) {
            continue;
        }

        /*
         * Il proxy deve poter utilizzare il cookie su tutte le sue URL.
         */
        if (str_starts_with(
            $lower,
            'path='
        )) {
            continue;
        }

        /*
         * Il cookie viene inviato anche in HTTP se il proxy viene usato
         * in HTTP.
         */
        if ($lower === 'secure') {
            continue;
        }

        /*
         * Evita problemi di policy del browser.
         */
        if (str_starts_with(
            $lower,
            'samesite='
        )) {
            continue;
        }

        /*
         * HttpOnly può essere mantenuto.
         */
        $result[] = $part;
    }

    /*
     * Il cookie deve valere per il proxy.
     */
    $result[] = 'Path=/';

    return implode(
        '; ',
        $result
    );
}


/**
 * Estrae tutti i cookie ricevuti dal server remoto.
 */
function extractSetCookies(
    array $headers
): array {

    $cookies = [];

    foreach ($headers as $header) {

        if (
            stripos(
                $header,
                'Set-Cookie:'
            ) !== 0
        ) {
            continue;
        }

        $value = trim(
            substr(
                $header,
                strlen('Set-Cookie:')
            )
        );

        if ($value !== '') {
            $cookies[] =
                rewriteSetCookieForProxy(
                    $value
                );
        }
    }

    return $cookies;
}


/*
 * ============================================================================
 * DOWNLOAD / PROXY DELLA RISORSA
 * ============================================================================
 */

/**
 * Scarica la risorsa remota.
 *
 * $followRedirects:
 *     true  = cURL segue i redirect.
 *     false = il redirect viene restituito al proxy.
 */
function fetchRemote(
    string $url,
    string $method,
    array $getData,
    array $postData,
    bool $followRedirects
): array {

    /*
     * ------------------------------------------------------------------------
     * QUERY STRING
     * ------------------------------------------------------------------------
     *
     * "url" è il parametro interno del proxy.
     *
     * Tutti gli altri parametri GET vengono inoltrati al server remoto.
     */

    if (
        $method === 'GET' &&
        !empty($getData)
    ) {

        $remoteGetData = $getData;

        unset(
            $remoteGetData['url']
        );

        if (!empty($remoteGetData)) {

            $query = http_build_query(
                $remoteGetData,
                '',
                '&',
                PHP_QUERY_RFC3986
            );

            if ($query !== '') {

                $separator =
                    str_contains($url, '?')
                    ? '&'
                    : '?';

                $url .=
                    $separator .
                    $query;
            }
        }
    }


    $responseHeaders = [];


    $ch = curl_init($url);


    $options = [

        CURLOPT_RETURNTRANSFER => true,

        CURLOPT_FOLLOWLOCATION =>
            $followRedirects,

        CURLOPT_MAXREDIRS => 5,

        CURLOPT_CONNECTTIMEOUT => 10,

        CURLOPT_TIMEOUT => 30,

        CURLOPT_USERAGENT =>
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' .
            'AppleWebKit/537.36 (KHTML, like Gecko) ' .
            'Chrome/153.0.0.0 Safari/537.36',

        CURLOPT_HTTPHEADER => [
            'Accept: */*',
        ],

        /*
         * Riceve gli header HTTP.
         */
        CURLOPT_HEADERFUNCTION =>
            function (
                $curl,
                string $header
            ) use (
                &$responseHeaders
            ) {

                $responseHeaders[] =
                    rtrim($header, "\r\n");

                return strlen($header);
            },
    ];


    /*
     * ------------------------------------------------------------------------
     * COOKIE DAL BROWSER
     * ------------------------------------------------------------------------
     */

    if (!empty($_COOKIE)) {

        $cookiePairs = [];

        foreach (
            $_COOKIE as $name => $value
        ) {

            $cookiePairs[] =
                $name . '=' . $value;
        }

        if (!empty($cookiePairs)) {

            $options[CURLOPT_COOKIE] =
                implode(
                    '; ',
                    $cookiePairs
                );
        }
    }


    /*
     * ------------------------------------------------------------------------
     * POST
     * ------------------------------------------------------------------------
     */

    if ($method === 'POST') {

        $options[CURLOPT_POST] = true;

        $options[CURLOPT_POSTFIELDS] =
            http_build_query(
                $postData,
                '',
                '&',
                PHP_QUERY_RFC3986
            );

        $options[CURLOPT_HTTPHEADER][] =
            'Content-Type: application/x-www-form-urlencoded';
    }


    /*
     * ------------------------------------------------------------------------
     * GET
     * ------------------------------------------------------------------------
     */

    else {

        $options[CURLOPT_HTTPGET] = true;
    }


    curl_setopt_array(
        $ch,
        $options
    );


    $body = curl_exec($ch);


    if ($body === false) {

        $error = curl_error($ch);

        curl_close($ch);

        return [
            'success' => false,
            'error'   => $error,
        ];
    }


    $status = curl_getinfo(
        $ch,
        CURLINFO_HTTP_CODE
    );


    $contentType = curl_getinfo(
        $ch,
        CURLINFO_CONTENT_TYPE
    );


    $finalUrl = curl_getinfo(
        $ch,
        CURLINFO_EFFECTIVE_URL
    );


    /*
     * Se FOLLOWLOCATION è false, CURLINFO_REDIRECT_URL
     * contiene l'URL del redirect.
     */
    $redirectUrl = curl_getinfo(
        $ch,
        CURLINFO_REDIRECT_URL
    );


    curl_close($ch);


    return [

        'success' => true,

        'status' => $status,

        'contentType' =>
            $contentType ?: '',

        'finalUrl' =>
            $finalUrl ?: $url,

        'redirectUrl' =>
            $redirectUrl ?: '',

        'headers' =>
            $responseHeaders,

        'body' => $body,
    ];
}


/*
 * ============================================================================
 * URL RICHIESTO
 * ============================================================================
 */

$url = $_GET['url'] ?? '';


if ($url === '') {
    $url = $defaultUrl;
}


/*
 * Decodifica il parametro.
 */
$url = urldecode(
    $url
);


/*
 * Sicurezza:
 * permettiamo solamente www.cartelli.it
 */
if (!isAllowedUrl(
    $url,
    $allowedHost
)) {

    http_response_code(403);

    header(
        'Content-Type: text/plain; charset=UTF-8'
    );

    exit(
        'URL non autorizzato.'
    );
}


/*
 * ============================================================================
 * RICHIESTA ORIGINALE DEL BROWSER
 * ============================================================================
 */

$method = strtoupper(
    $_SERVER['REQUEST_METHOD'] ?? 'GET'
);


/*
 * ============================================================================
 * RICHIESTA PRINCIPALE VS RISORSA
 * ============================================================================
 *
 * Per la richiesta principale vogliamo intercettare i redirect del server
 * remoto e restituirli al browser.
 *
 * Per le risorse secondarie possiamo invece seguire normalmente i redirect.
 *
 * Nel nostro caso consideriamo richiesta principale una richiesta HTML
 * effettuata tramite catalogo.php senza il parametro interno "resource".
 */

$isMainRequest =
    !isset($_GET['resource']) ||
    $_GET['resource'] !== '1';


/*
 * POST originale.
 */
$postData = $_POST;


/*
 * ============================================================================
 * SCARICA RISORSA
 * ============================================================================
 */

$response = fetchRemote(
    $url,
    $method,
    $_GET,
    $postData,
    !$isMainRequest
);


if (!$response['success']) {

    http_response_code(502);

    header(
        'Content-Type: text/plain; charset=UTF-8'
    );

    echo 'Errore proxy: ' .
         $response['error'];

    exit;
}


/*
 * ============================================================================
 * GESTIONE REDIRECT DELLA PAGINA PRINCIPALE
 * ============================================================================
 */

if (
    $isMainRequest &&
    in_array(
        $response['status'],
        [301, 302, 303, 307, 308],
        true
    )
) {

    $redirectUrl =
        $response['redirectUrl'];


    /*
     * Se CURLINFO_REDIRECT_URL non è disponibile,
     * proviamo a ricavare Location dagli header.
     */
    if ($redirectUrl === '') {

        foreach (
            $response['headers']
            as $header
        ) {

            if (
                stripos(
                    $header,
                    'Location:'
                ) === 0
            ) {

                $redirectUrl = trim(
                    substr(
                        $header,
                        strlen('Location:')
                    )
                );

                break;
            }
        }
    }


    if ($redirectUrl !== '') {

        /*
         * Il Location può essere relativo.
         */
        $redirectUrl = absoluteUrl(
            $redirectUrl,
            $url
        );


        /*
         * Sicurezza:
         * il redirect deve rimanere sul dominio autorizzato.
         */
        if (!isAllowedUrl(
            $redirectUrl,
            $allowedHost
        )) {

            http_response_code(502);

            header(
                'Content-Type: text/plain; charset=UTF-8'
            );

            exit(
                'Redirect verso dominio non autorizzato.'
            );
        }


        /*
         * Salva eventuali cookie restituiti dal server remoto.
         */
        $setCookies =
            extractSetCookies(
                $response['headers']
            );

        foreach (
            $setCookies as $cookie
        ) {

            header(
                'Set-Cookie: ' . $cookie,
                false
            );
        }


        /*
         * Trasforma il Location remoto in URL del proxy.
         */
        $proxyRedirect =
            proxyUrl($redirectUrl);


        /*
         * Per un 303 il browser deve effettuare GET.
         *
         * Per gli altri status manteniamo il codice originale.
         */
        http_response_code(
            $response['status']
        );


        header(
            'Location: ' .
            $proxyRedirect
        );


        exit;
    }
}


/*
 * ============================================================================
 * STATUS HTTP
 * ============================================================================
 */

if (
    $response['status'] < 200 ||
    $response['status'] >= 400
) {

    http_response_code(
        $response['status'] ?: 502
    );


    if (
        $response['contentType'] !== ''
    ) {

        header(
            'Content-Type: ' .
            $response['contentType']
        );
    }


    echo $response['body'];

    exit;
}


$contentType = strtolower(
    explode(
        ';',
        $response['contentType']
    )[0]
);


$finalUrl =
    $response['finalUrl'];


$body =
    $response['body'];


/*
 * ============================================================================
 * COOKIE DELLA RISPOSTA REMOTA
 * ============================================================================
 */

$setCookies =
    extractSetCookies(
        $response['headers']
    );

foreach (
    $setCookies as $cookie
) {

    header(
        'Set-Cookie: ' . $cookie,
        false
    );
}


/*
 * ============================================================================
 * CSS
 * ============================================================================
 */

if ($contentType === 'text/css') {


    /*
     * ------------------------------------------------------------------------
     * url(...)
     * ------------------------------------------------------------------------
     */

    $body = preg_replace_callback(

        '~url\(\s*([\'"]?)(.*?)\1\s*\)~i',

        function (array $match)
            use (
                $finalUrl,
                $allowedHost
            ) {

                $quote = $match[1];

                $resource = trim(
                    $match[2]
                );

                if (
                    $resource === '' ||
                    str_starts_with(
                        $resource,
                        '#'
                    ) ||
                    preg_match(
                        '~^(data|blob|javascript):~i',
                        $resource
                    )
                ) {
                    return $match[0];
                }

                $absolute = absoluteUrl(
                    $resource,
                    $finalUrl
                );

                if (!isAllowedUrl(
                    $absolute,
                    $allowedHost
                )) {
                    return $match[0];
                }

                return 'url(' .
                       $quote .
                       proxyUrl($absolute) .
                       $quote .
                       ')';
            },

        $body
    );


    /*
     * ------------------------------------------------------------------------
     * @import
     * ------------------------------------------------------------------------
     */

    $body = preg_replace_callback(

        '~@import\s+(?:url\(\s*)?[\'"]?([^\'"\)\s]+)[\'"]?\s*\)?~i',

        function (array $match)
            use (
                $finalUrl,
                $allowedHost
            ) {

                $resource = trim(
                    $match[1]
                );

                $absolute = absoluteUrl(
                    $resource,
                    $finalUrl
                );

                if (!isAllowedUrl(
                    $absolute,
                    $allowedHost
                )) {
                    return $match[0];
                }

                return '@import url("' .
                       proxyUrl($absolute) .
                       '")';
            },

        $body
    );


    header(
        'Content-Type: text/css; charset=UTF-8'
    );

    header(
        'Cache-Control: public, max-age=3600'
    );

    echo $body;

    exit;
}


/*
 * ============================================================================
 * HTML
 * ============================================================================
 */

if (
    $contentType === 'text/html' ||
    $contentType === 'application/xhtml+xml' ||
    $contentType === ''
) {

    libxml_use_internal_errors(true);

    $dom = new DOMDocument();

    $dom->loadHTML(
        $body,
        LIBXML_HTML_NOIMPLIED |
        LIBXML_HTML_NODEFDTD
    );

    $xpath = new DOMXPath($dom);


    /*
     * ------------------------------------------------------------------------
     * RIMOZIONE ELEMENTI
     * ------------------------------------------------------------------------
     */

    $selectors = [

        '//*[@id="advertisement"]',

        '//*[@class="cookie-banner"]',

        '//header',

        '//span[contains(@id, "prodotto_lblDescPrezzoListino")]', 

        '//span[contains(@id, "prodotto_lblPrezzoListino")]', 

        '//span[contains(@id, "ctl00_cphGeneralMasterPage_griglia_ctl01_lblErrore")]', 
    ];


    foreach (
        $selectors as $selector
    ) {

        foreach (
            $xpath->query($selector)
            as $node
        ) {

            $node->parentNode?->removeChild(
                $node
            );
        }
    }


    /*
     * ------------------------------------------------------------------------
     * LINK <a href="">
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//a[@href]')
        as $node
    ) {

        $node->setAttribute(

            'href',

            rewriteUrl(

                $node->getAttribute(
                    'href'
                ),

                $finalUrl,

                $allowedHost
            )
        );
    }


    /*
     * ------------------------------------------------------------------------
     * FORM <form action="">
     * ------------------------------------------------------------------------
     *
     * Se action è vuoto, il browser utilizza normalmente l'URL corrente.
     *
     * In quel caso usiamo $finalUrl.
     */

    foreach (
        $xpath->query('//form')
        as $form
    ) {

        $action = trim(
            $form->getAttribute(
                'action'
            )
        );

        if ($action === '') {

            $action = $finalUrl;
        }

        $form->setAttribute(

            'action',

            rewriteUrl(

                $action,

                $finalUrl,

                $allowedHost
            )
        );
    }


    /*
     * ------------------------------------------------------------------------
     * formaction
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//*[@formaction]')
        as $element
    ) {

        $formAction = trim(
            $element->getAttribute(
                'formaction'
            )
        );

        if ($formAction === '') {
            continue;
        }

        $element->setAttribute(

            'formaction',

            rewriteUrl(

                $formAction,

                $finalUrl,

                $allowedHost
            )
        );
    }


    /*
     * ------------------------------------------------------------------------
     * TUTTI GLI ELEMENTI CON src
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//*[@src]')
        as $node
    ) {

        $src = trim(
            $node->getAttribute('src')
        );

        if ($src === '') {
            continue;
        }

        $node->setAttribute(

            'src',

            rewriteUrl(

                $src,

                $finalUrl,

                $allowedHost
            )
        );
    }


    /*
     * ------------------------------------------------------------------------
     * LINK <link href="">
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//link[@href]')
        as $node
    ) {

        $href = trim(
            $node->getAttribute('href')
        );

        if ($href === '') {
            continue;
        }

        $node->setAttribute(

            'href',

            rewriteUrl(

                $href,

                $finalUrl,

                $allowedHost
            )
        );
    }


    /*
     * ------------------------------------------------------------------------
     * srcset
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//*[@srcset]')
        as $node
    ) {

        $srcset =
            $node->getAttribute(
                'srcset'
            );

        $items =
            explode(
                ',',
                $srcset
            );

        $newItems = [];


        foreach (
            $items as $item
        ) {

            $item = trim(
                $item
            );

            if ($item === '') {
                continue;
            }

            $parts = preg_split(
                '/\s+/',
                $item,
                2
            );

            $resource =
                $parts[0];

            $descriptor =
                $parts[1] ?? '';


            $resource =
                rewriteUrl(

                    $resource,

                    $finalUrl,

                    $allowedHost
                );


            $newItems[] =
                $resource .
                (
                    $descriptor !== ''
                    ? ' ' . $descriptor
                    : ''
                );
        }


        $node->setAttribute(

            'srcset',

            implode(
                ', ',
                $newItems
            )
        );
    }


    /*
     * ------------------------------------------------------------------------
     * CSS INLINE
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//*[@style]')
        as $node
    ) {

        $style =
            $node->getAttribute(
                'style'
            );


        $style =
            preg_replace_callback(

                '~url\(\s*([\'"]?)(.*?)\1\s*\)~i',

                function (array $match)
                    use (
                        $finalUrl,
                        $allowedHost
                    ) {

                        $quote =
                            $match[1];

                        $resource =
                            trim(
                                $match[2]
                            );


                        if (
                            $resource === '' ||
                            preg_match(
                                '~^(data|blob|javascript):~i',
                                $resource
                            )
                        ) {
                            return $match[0];
                        }


                        $absolute =
                            absoluteUrl(
                                $resource,
                                $finalUrl
                            );


                        if (!isAllowedUrl(
                            $absolute,
                            $allowedHost
                        )) {
                            return $match[0];
                        }


                        return 'url(' .
                               $quote .
                               proxyUrl($absolute) .
                               $quote .
                               ')';
                    },

                $style
            );


        $node->setAttribute(
            'style',
            $style
        );
    }


    /*
     * ------------------------------------------------------------------------
     * RIMUOVI <base>
     * ------------------------------------------------------------------------
     */

    foreach (
        $xpath->query('//base')
        as $base
    ) {

        $base->parentNode?->removeChild(
            $base
        );
    }


    /*
     * ------------------------------------------------------------------------
     * CSS PERSONALIZZATO DEL PROXY
     * ------------------------------------------------------------------------
     */

    $style = $dom->createElement(
        'style'
    );


    $style->appendChild(
        $dom->createTextNode(
            '#sidebarDx,
             .footer,
             #secondary-menu-bar,
             #main-menu-top,
             #Cataloghi,
             .button.btn-registrati {
                display: none !important;
             }

             div.contenitore {
                margin-top: 0 !important;
             }'
        )
    );


    $head =
        $xpath->query('//head')->item(0);


    if ($head !== null) {

        $head->appendChild(
            $style
        );

    } else {

        $dom->documentElement->insertBefore(
            $style,
            $dom->documentElement->firstChild
        );
    }


    /*
     * ------------------------------------------------------------------------
     * OUTPUT HTML
     * ------------------------------------------------------------------------
 */

    header(
        'Content-Type: text/html; charset=UTF-8'
    );


    echo $dom->saveHTML();

    exit;
}


/*
 * ============================================================================
 * TUTTE LE ALTRE RISORSE
 * ============================================================================
 *
 * Immagini, font, JS, JSON, SVG, ecc.
 * ============================================================================
 */

if ($contentType !== '') {

    header(
        'Content-Type: ' .
        $contentType
    );
}


header(
    'Cache-Control: public, max-age=3600'
);


echo $body;
