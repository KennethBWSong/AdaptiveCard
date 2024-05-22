const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  loginStartPage: process.env.INITIATE_LOGIN_ENDPOINT,
  clientId: process.env.M365_CLIENT_ID,
  clientSecret: process.env.M365_CLIENT_SECRET,
  tenantId: process.env.M365_TENANT_ID,
  appIdUri: process.env.M365_APPLICATION_ID_URI
};

export default config;
