// pages.dev本番ホストへのアクセスを正規ドメインに301リダイレクトする
// （_redirectsはドメイン単位のリダイレクト非対応のためFunctionsで実装）
// プレビューURL（*.travel-english-blog.pages.dev）は動作確認用に除外する
export async function onRequest(context) {
    const url = new URL(context.request.url);
    if (url.hostname === 'travel-english-blog.pages.dev') {
        url.hostname = 'tabisuru-eigo.com';
        return Response.redirect(url.toString(), 301);
    }
    return context.next();
}
