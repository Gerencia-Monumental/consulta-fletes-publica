const axios = require("axios");

module.exports = async function (context, req) {

    const plate = (req.body?.plate || "").toUpperCase();

    if (!plate) {
        context.res = {
            status: 400,
            body: "Placa requerida"
        };
        return;
    }

    const tokenResp = await axios.post(
        `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
        new URLSearchParams({
            client_id: process.env.CLIENT_ID,
            client_secret: process.env.CLIENT_SECRET,
            scope: "https://graph.microsoft.com/.default",
            grant_type: "client_credentials"
        }),
        { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    const token = tokenResp.data.access_token;

    const url =
        `https://graph.microsoft.com/v1.0/drives/${process.env.DRIVE_ID}` +
        `/items/${process.env.ITEM_ID}/workbook/tables('${process.env.TABLE_NAME}')/range`;

    const resp = await axios.get(url, {
        headers: { Authorization: `Bearer ${token}` }
    });

    const values = resp.data.values;

    const headers = values[0];
    const rows = values.slice(1);

    const placaIndex = headers.indexOf("PLACA");

    const matches = rows.filter(r =>
        (r[placaIndex] || "").toUpperCase() === plate
    );

    context.res = {
        status: 200,
        body: matches
    };
};
