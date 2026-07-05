function round2(n) {
    return Math.round((Number(n) || 0) * 100) / 100;
}

function buildDateTextKey(dateValue, textValue) {
    const dateKey = new Date(dateValue).toISOString().slice(0, 10);
    return `${dateKey}|${textValue}`;
}
