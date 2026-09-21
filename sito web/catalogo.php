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
 * DOWNLOAD / PROXY DELLA RISORSA
 * ============================================================================
 */

/**
 * Scarica la risorsa remota.
 *
 * Per POST ricostruisce il body a partire da $_POST.
 */
function fetchRemote(
    string $url,
    string $method,
    array $postData
): array {

    $ch = curl_init($url);

    $options = [

        CURLOPT_RETURNTRANSFER => true,

        CURLOPT_FOLLOWLOCATION => true,

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
    ];


    /*
     * ------------------------------------------------------------------------
     * POST
     * ------------------------------------------------------------------------
     */

    if ($method === 'POST') {

        $options[CURLOPT_POST] = true;

        /*
         * Ricostruisce il POST a partire da $_POST.
         */
        $options[CURLOPT_POSTFIELDS] = http_build_query(
            $postData,
            '',
            '&'
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


    curl_close($ch);


    return [

        'success' => true,

        'status' => $status,

        'contentType' =>
            $contentType ?: '',

        'finalUrl' =>
            $finalUrl ?: $url,

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
$url = urldecode($url);


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
 * Ricostruzione del POST esclusivamente da $_POST.
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
    $postData
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

    if ($response['contentType'] !== '') {

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

    ];

    foreach ($selectors as $selector) {

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
     */

    foreach (
        $xpath->query('//form[@action]')
        as $form
    ) {

        $action = trim(
            $form->getAttribute('action')
        );

        if ($action === '') {
            continue;
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
     *
     * Gestisce:
     *
     * <button formaction="...">
     * <input formaction="...">
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
     *
     * Comprende automaticamente:
     *
     * <img src="">
     * <script src="">
     * <iframe src="">
     * <video src="">
     * <audio src="">
     * <source src="">
     * <embed src="">
     * <input type="image" src="">
     * ecc.
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
     *
     * CSS, favicon, preload, ecc.
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

        $srcset = $node->getAttribute(
            'srcset'
        );

        $items = explode(
            ',',
            $srcset
        );

        $newItems = [];

        foreach ($items as $item) {

            $item = trim($item);

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

            $resource = rewriteUrl(

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

        $style = $node->getAttribute(
            'style'
        );

        $style = preg_replace_callback(

            '~url\(\s*([\'"]?)(.*?)\1\s*\)~i',

            function (array $match)
                use (
                    $finalUrl,
                    $allowedHost
                ) {

                    $quote =
                        $match[1];

                    $resource =
                        trim($match[2]);

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
 *
 * Vengono restituite senza modifiche.
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
