// Static-asset requests (version.txt, index.html, manual_version.txt, the
// manual docx) never reach this handler -- Cloudflare serves them directly
// from the [assets] directory configured in wrangler.toml. This only runs
// for a request that matches no file at all.
export default {
	async fetch() {
		return new Response("Not Found", { status: 404 });
	},
};
