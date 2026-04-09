import {
  ExchangeService,
  ExchangeVersion,
  WebCredentials,
  Uri,
  WellKnownFolderName,
  ItemView
} from "ews-javascript-api";

import * as dotenv from "dotenv";
dotenv.config();

async function main() {
  const email = process.env.EWS_EMAIL;
  const password = process.env.EWS_PASSWORD;
  const url = process.env.EWS_URL || "https://email.metib.ru/EWS/Exchange.asmx";

  if (!email || !password) {
    throw new Error("Missing EWS_EMAIL or EWS_PASSWORD in env");
  }

  const service = new ExchangeService(ExchangeVersion.Exchange2016);

  service.Credentials = new WebCredentials(email, password);
  service.Url = new Uri(url);

  try {
    const res = await service.FindItems(
      WellKnownFolderName.Drafts,
      new ItemView(5)
    );

    console.log("✅ Connected");
    console.log("Drafts:", res.TotalCount);

    for (const item of res.Items) {
      console.log("-", item.Subject);
    }

  } catch (e) {
    console.error("❌ Error:", e);
  }
}

main();
