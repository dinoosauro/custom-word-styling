interface Props {
    /**
     * The function that'll be called if it's not possible to automatically open the user's browser
     * @param url the URL to download the content
     */
    urlCallback: (url: string) => void,
    /**
     * An array of bool values, that indicate if the style at that position should be exported or not
     */
    propertiesToExport: boolean[]
}

/**
 * Export the selected styles in a JSON file
 */
export default function ExportStyles({propertiesToExport, urlCallback}: Props) {
    const spinner = document.createElement("div");
    spinner.classList.add("spinner");
    document.body.append(spinner);
    setTimeout(async () => {
        await Word.run(async (ctx) => {
            /**
             * Load all the properties of all the styles
             */
            const styles = ctx.document.getStyles().load({
                // @ts-ignore
                $all: true,
                borders: true,
                font: true,
                tableStyle: true,
                listTemplate: true,
                paragraphFormat: true,
                shading: true,
            });
            await ctx.sync();
            // Since downloads aren't supported in the taskpane, we'll open the `downloader.html` page in the default browser to download the content. The filename and the data will be added in the page hash
            const url = new URL(window.location.href);
            url.pathname = `${url.pathname.substring(0, url.pathname.lastIndexOf("/"))}/pathname`;
            let possibleFileName = Office.context.document.url
                ?.split("/")
                .pop();
            if (typeof possibleFileName !== "undefined")
                possibleFileName = possibleFileName.substring(
                    0,
                    possibleFileName.lastIndexOf("."),
                );
            url.hash = new URLSearchParams({
                data: JSON.stringify({
                    styles: styles.items.filter((a, i) => propertiesToExport[i]),
                }),
                name: `Styles-${possibleFileName ?? Date.now()}.json`,
            }).toString();
            if (Office.isSetSupported("OpenBrowserWindowApi", "1.1")) {
                Office.context.ui.openBrowserWindow(url.toString());
            } else urlCallback(url.toString()); // Fallback: a dialog will be displayed with the link to download it.
        });
        spinner.remove();
    }, 1);

}