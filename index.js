export default {
  async fetch(request, env, ctx) {
    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    };

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders });
    }

    const url = new URL(request.url);
    const email = url.searchParams.get("email");

    if (!email) {
      return new Response(JSON.stringify({ error: "Email manquant" }), { status: 400, headers: corsHeaders });
    }

    try {
      const tokenParams = new URLSearchParams({
        client_id: env.CLIENT_ID,
        scope: 'https://graph.microsoft.com/.default',
        client_secret: env.CLIENT_SECRET,
        grant_type: 'client_credentials'
      });

      const tokenResponse = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: tokenParams
      });

      if (!tokenResponse.ok) throw new Error("Échec authentification Microsoft");

      const accessToken = (await tokenResponse.json()).access_token;
      const graphUrl = `https://graph.microsoft.com/v1.0/users/${email}?$select=displayName,jobTitle,mobilePhone,businessPhones`;
      
      const graphResponse = await fetch(graphUrl, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });

      if (!graphResponse.ok) {
         return new Response(JSON.stringify({ error: "Utilisateur introuvable" }), { status: 404, headers: corsHeaders });
      }

      const userData = await graphResponse.json();

      // Extraction conditionnelle stricte des variables
      const nom = userData.displayName || "Collaborateur";
      const posteHtml = userData.jobTitle ? `<p style="margin: 0; color: #666;">${userData.jobTitle}</p>` : "";
      
      let telHtml = "";
      if (userData.mobilePhone) {
          telHtml = `<p style="margin: 0;">Tel : ${userData.mobilePhone}</p>`;
      } else if (userData.businessPhones && userData.businessPhones.length > 0) {
          telHtml = `<p style="margin: 0;">Tel : ${userData.businessPhones[0]}</p>`;
      }

      const currentBanner = "https://solacomete.github.io/signature-m365/icon-128.png";

      // Assemblage modulaire du HTML final
      const htmlSignature = `
        <br><br>
        <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #333;">
            <p style="margin: 0;"><strong>${nom}</strong></p>
            ${posteHtml}
            <p style="margin: 0;">${email}</p>
            ${telHtml}
            <br>
            <img src="${currentBanner}" alt="Actualité Entreprise" style="max-width: 400px;">
        </div>
      `;

      return new Response(JSON.stringify({ html: htmlSignature }), {
        headers: { "Content-Type": "application/json", ...corsHeaders }
      });

    } catch (error) {
      return new Response(JSON.stringify({ error: "Erreur serveur" }), { status: 500, headers: corsHeaders });
    }
  }
};
