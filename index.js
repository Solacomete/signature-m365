export default {
  async fetch(request, env, ctx) {
    // 1. Autorisations CORS pour Outlook
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
      // 2. Demande du jeton d'accès (Token) à Microsoft Entra ID
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

      if (!tokenResponse.ok) {
        throw new Error("Échec de l'authentification Microsoft");
      }

      const tokenData = await tokenResponse.json();
      const accessToken = tokenData.access_token;

      // 3. Interrogation de Microsoft Graph pour récupérer l'utilisateur
      const graphUrl = `https://graph.microsoft.com/v1.0/users/${email}?$select=displayName,jobTitle,mobilePhone,businessPhones`;
      const graphResponse = await fetch(graphUrl, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });

      if (!graphResponse.ok) {
         return new Response(JSON.stringify({ error: "Utilisateur introuvable dans l'annuaire" }), { status: 404, headers: corsHeaders });
      }

      const userData = await graphResponse.json();

      // 4. Extraction des données (avec gestion des champs vides si le profil n'est pas rempli côté RH)
      const nom = userData.displayName || "Collaborateur";
      const poste = userData.jobTitle || " ";
      // On prend le mobile en priorité, sinon le tel fixe (businessPhones est un tableau)
      const tel = userData.mobilePhone || (userData.businessPhones && userData.businessPhones.length > 0 ? userData.businessPhones[0] : " ");

      // L'URL de ta bannière actuelle
      const currentBanner = "https://solacomete.github.io/signature-m365/icon-128.png";

      // 5. Génération de la signature finale
      const htmlSignature = `
        <br><br>
        <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #333;">
            <p style="margin: 0;"><strong>${nom}</strong></p>
            <p style="margin: 0; color: #666;">${poste}</p>
            ${tel !== " " ? `<p style="margin: 0;">Tel : ${tel}</p>` : ""}
            <br>
            <img src="${currentBanner}" alt="Actualité Entreprise" style="max-width: 400px;">
        </div>
      `;

      return new Response(JSON.stringify({ html: htmlSignature }), {
        headers: { "Content-Type": "application/json", ...corsHeaders }
      });

    } catch (error) {
      return new Response(JSON.stringify({ error: "Erreur serveur interne" }), { status: 500, headers: corsHeaders });
    }
  }
};
