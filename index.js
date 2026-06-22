export default {
  async fetch(request, env, ctx) {
    // 1. Autoriser Outlook Web à interroger cette API (Règles CORS obligatoires)
    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    // Réponse de sécurité "Preflight" demandée par les navigateurs
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders });
    }

    // 2. Analyser la requête d'Outlook (ex: ?email=louis@dagobertindustrie.com)
    const url = new URL(request.url);
    const email = url.searchParams.get("email");

    // 3. Ta base de données interne (À compléter avec tes collaborateurs)
    const usersDB = {
      "louis@dagobertindustrie.com": {
        nom: "Louis",
        fonction: "Responsable SI",
        tel: "06 11 22 33 44"
      },
      "j.nesta@apisoudure.com": {
        nom: "Jules NESTA",
        fonction: "Commercial",
        tel: "06 99 88 77 66"
      }
    };

    // 4. L'URL de la bannière active (à modifier ici lors des événements)
    const currentBanner = "https://solacomete.github.io/signature-m365/icon-128.png"; // Image de test

    // Si l'email n'est pas dans la base
    if (!email || !usersDB[email]) {
      return new Response(JSON.stringify({ error: "Utilisateur non trouvé" }), {
        status: 404,
        headers: { "Content-Type": "application/json", ...corsHeaders }
      });
    }

    // 5. Génération du HTML avec les vraies données
    const user = usersDB[email];
    const htmlSignature = `
      <br><br>
      <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #333;">
          <p style="margin: 0;"><strong>${user.nom}</strong></p>
          <p style="margin: 0; color: #666;">${user.fonction}</p>
          <p style="margin: 0;">Tel : ${user.tel}</p>
          <br>
          <img src="${currentBanner}" alt="Actualité Entreprise" style="max-width: 400px;">
      </div>
    `;

    // 6. Envoi de la signature à Outlook
    return new Response(JSON.stringify({ html: htmlSignature }), {
      headers: { "Content-Type": "application/json", ...corsHeaders }
    });
  }
};
