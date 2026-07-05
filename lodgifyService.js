function getLodgifyConfig() {
    const props = PropertiesService.getScriptProperties();
    const apiKey = (props.getProperty("LODGIFY_API_KEY") || "").trim();
    const baseUrl = (props.getProperty("LODGIFY_API_BASE_URL") || "https://api.lodgify.com/v1").trim();
    const timeoutMs = Number(props.getProperty("LODGIFY_TIMEOUT_MS") || 30000);

    return {
        apiKey,
        baseUrl: baseUrl.replace(/\/+$/, ""),
        timeoutMs: Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 30000
    };
}

function validateLodgifyConfig() {
    const config = getLodgifyConfig();
    if (!config.apiKey) {
        throw new Error("Missing Script Property: LODGIFY_API_KEY");
    }
    if (!config.baseUrl) {
        throw new Error("Missing Script Property: LODGIFY_API_BASE_URL");
    }
    return config;
}

function setLodgifyConfig(apiKey, baseUrl, timeoutMs) {
    if (!apiKey || !String(apiKey).trim()) {
        throw new Error("apiKey must not be empty");
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperty("LODGIFY_API_KEY", String(apiKey).trim());

    if (baseUrl && String(baseUrl).trim()) {
        props.setProperty("LODGIFY_API_BASE_URL", String(baseUrl).trim());
    }

    if (timeoutMs !== undefined && timeoutMs !== null) {
        props.setProperty("LODGIFY_TIMEOUT_MS", String(timeoutMs));
    }
}

function lodgifyBuildUrl(path, queryParams) {
    const config = validateLodgifyConfig();
    const normalizedPath = path.startsWith("/") ? path : `/${path}`;
    const base = `${config.baseUrl}${normalizedPath}`;

    if (!queryParams) return base;

    const parts = [];
    Object.keys(queryParams).forEach(key => {
        const value = queryParams[key];
        if (value === undefined || value === null || value === "") return;
        parts.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
    });

    if (parts.length === 0) return base;
    return `${base}?${parts.join("&")}`;
}

function lodgifyRequest(path, options) {
    const config = validateLodgifyConfig();
    const requestOptions = options || {};
    const method = (requestOptions.method || "get").toLowerCase();

    const fetchOptions = {
        method,
        muteHttpExceptions: true,
        headers: {
            "X-ApiKey": config.apiKey,
            Accept: "application/json"
        }
    };

    if (requestOptions.payload !== undefined && requestOptions.payload !== null) {
        fetchOptions.contentType = "application/json";
        fetchOptions.payload = JSON.stringify(requestOptions.payload);
    }

    const url = lodgifyBuildUrl(path, requestOptions.queryParams);
    const response = UrlFetchApp.fetch(url, fetchOptions);
    const status = response.getResponseCode();
    const bodyText = response.getContentText() || "";

    let body;
    try {
        body = bodyText ? JSON.parse(bodyText) : null;
    } catch (e) {
        body = bodyText;
    }

    if (status < 200 || status >= 300) {
        throw new Error(`Lodgify request failed (${status}) for ${url}: ${bodyText}`);
    }

    return {
        status,
        body,
        headers: response.getAllHeaders()
    };
}

function lodgifyGetBookings(queryParams) {
    return lodgifyRequest("/bookings", {
        method: "get",
        queryParams: queryParams || {}
    });
}

function lodgifyGetReservations(queryParams) {
    return lodgifyRequest("/reservations", {
        method: "get",
        queryParams: queryParams || {}
    });
}

function lodgifyHealthCheck() {
    return lodgifyGetBookings({ page: 1, size: 1 });
}
