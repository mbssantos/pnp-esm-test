import { Configuration } from "@azure/msal-node";
import { SPDefault } from "@pnp/nodejs";
import { SPFI, spfi } from "@pnp/sp";

// // works with module=CommonJS AND moduleResolution=node10
// // Doesn't work with NodeNext (as expected)
// import "@pnp/sp/webs";

// Doesn't work with NodeNext (unexpected)
import "@pnp/sp/webs/index.js";
import "@pnp/sp/webs/types.js";

const createClient = (site: string) => {
  const config: Configuration = {
    auth: {
      authority: "AUTHORITY",
      clientCertificate: {
        thumbprint: "THUMBPRINT",
        privateKey: "PRIVATE_KEY",
      },
      clientId: "AZURE_CLIENT_ID",
      knownAuthorities: ["AUTHORITY"],
    },
  };

  const client = spfi().using(
    SPDefault({
      baseUrl: `https://AZURE_TENANT.sharepoint.com/sites/${site}`,
      msal: {
        config: config,
        scopes: [`https://AZURE_TENANT.sharepoint.com/.default`],
      },
    })
  );

  // TS complains in the next line: Property 'web' does not exist on type 'SPFI'
  console.log(client.web);

  // client.
  return client;
};

const clientPool: Record<string, SPFI> = {};

const SpClient = {
  getClient(site: string) {
    if (!clientPool[site]) {
      clientPool[site] = createClient(site);
    }

    return clientPool[site];
  },
};

export default SpClient;
