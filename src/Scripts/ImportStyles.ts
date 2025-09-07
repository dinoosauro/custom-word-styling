/**
 * Import styles by reading a JSON file from the user's drive
 * @param mergeStyleOption the behavior the script should keep if a style with the same name is found
 */
export default function ImportStyles(mergeStyleOption: "Overwrite" | "Ignore" | "CreateNew") {
    const spinner = document.createElement("div");
    spinner.classList.add("spinner");
    const input = Object.assign(document.createElement("input"), {
        type: "file",
        onchange: async () => {
            if (input.files) {
                // Theoretically, there's an API that should import all styles (`ctx.document.importStylesFromJson`). Practically, that is as useful as my life, and, since it doesn't work, we have to iterate over each element to update it.
                // And obviously, since Office.JS is made by Microsoft, we have to overcomplicate simple things. So, here is a script that'll handle in a semi-decent way the style changes.
                let json: { styles: Word.Style[] } = JSON.parse(
                    await input.files[0].text(),
                );
                document.body.append(spinner);
                setTimeout(async () => {
                    /**
                     * A list of all the styles that need to be created. This means that we'll need to run two different Word contexts: one to create them, and the other to update them.
                     * 
                     * The value of this object is composed of an array, that contains a string (the name of the new style) and the style fetched from the JSON file
                     */
                    let stylesThatShouldBeCreated: [string, Word.Style][] = [];
                    /**
                     * A list of all the styles that already exist in the file and must be edited.
                     * 
                     * The value of this object is composed of an array, that contains a number (the position in the `styles` array of the item to edit) and the style fetched from the JSON file
                     */
                    let itemsWithSyncAtTheEnd: [number, Word.Style][] = [];
                    /**
                     *  A list that contains [the original name of the style, and the name of the new style]. This is done so that, in case of duplicates, the `nextParagraphStyle` property can be edited.
                     */
                    let previousNames: [string, string][] = [];
                    /**
                     * Current progress of the apply process
                     */
                    let jsonProgress = 0;
                    await Word.run(async (ctx) => {
                        // @ts-ignore
                        const styles = ctx.document.getStyles().load(mergeStyleOption === "Overwrite" ? { $all: true, borders: true, paragraphFormat: true, font: true, listTemplate: true, shading: true, tableStyle: true } : { nameLocal: true });
                        await ctx.sync();
                        // Let's start by adding all the styles to their array
                        for (let i = 0; i < json.styles.length; i++) {
                            const isAlreadyThere = styles.items.findIndex(a => a.nameLocal === json.styles[i].nameLocal);
                            switch (mergeStyleOption) {
                                case "Overwrite":
                                    if (isAlreadyThere === -1) stylesThatShouldBeCreated.push([json.styles[i].nameLocal, json.styles[i]]); else itemsWithSyncAtTheEnd.push([isAlreadyThere, json.styles[i]]);
                                    break;
                                case "Ignore":
                                    if (isAlreadyThere === -1) stylesThatShouldBeCreated.push([json.styles[i].nameLocal, json.styles[i]]);
                                    break;
                                case "CreateNew":
                                    let suggestedName = json.styles[i].nameLocal;
                                    if (styles.items.findIndex(j => json.styles[i].nameLocal === j.nameLocal) !== -1) { // We need to get the first number that is not an entry 
                                        let duplicateIndex = 1;
                                        while (styles.items.findIndex(j => j.nameLocal === `${json.styles[i].nameLocal}${duplicateIndex}`) !== -1) duplicateIndex++;
                                        suggestedName = `${json.styles[i].nameLocal}${duplicateIndex}`;
                                        previousNames.push([json.styles[i].nameLocal, suggestedName]);
                                    }
                                    stylesThatShouldBeCreated.push([suggestedName, json.styles[i]]);
                                    break;
                            }
                        }
                        // We can start editing the properties that should be overwritten
                        for (const [stylePosition, jsonStyle] of itemsWithSyncAtTheEnd) {
                            document.body.style.setProperty(`--progress`, `${jsonProgress * 100 / json.styles.length * 3.6}deg`);
                            jsonProgress++;
                            await new Promise<void>((res) => {
                                setTimeout(async () => {
                                    await updateProperties(styles.items[stylePosition], jsonStyle);
                                    res();
                                }, 1)
                            })
                        }
                        // And we can also start creating the new properties that should be added
                        for (const [styleName, jsonUpdate] of stylesThatShouldBeCreated) ctx.document.addStyle(styleName, jsonUpdate.type);
                        await ctx.sync();
                    });
                    await new Promise(res => setTimeout(res, 150));
                    // Now, we'll edit the properties we've added
                    if (stylesThatShouldBeCreated.length > 0) {
                        await Word.run(async (ctx) => {
                            // @ts-ignore
                            const styles = ctx.document.getStyles().load({$all: true, borders: true, font: true, listTemplate: true, paragraphFormat: true, shading: true, tableStyle: true});
                            await ctx.sync();
                            for (const [nameOfNewStyle, jsonStyle] of stylesThatShouldBeCreated) {
                                // Update the `nextParagraphStyle` property if also that style has had its name changed.
                                const nextParagraphName = previousNames.find(i => i[0] === jsonStyle.nextParagraphStyle);
                                if (nextParagraphName) jsonStyle.nextParagraphStyle = nextParagraphName[1];
                                document.body.style.setProperty(`--progress`, `${jsonProgress * 100 / json.styles.length * 3.6}deg`);
                                jsonProgress++;
                                await new Promise<void>((res) => {
                                    setTimeout(async () => {
                                        const wordStyle = styles.items.find(i => i.nameLocal === nameOfNewStyle);
                                        if (!wordStyle) {
                                            res();
                                            return;
                                        }
                                        await ctx.sync();
                                        await updateProperties(wordStyle, jsonStyle);
                                        res();
                                    }, 1)
                                })
                            }
                            await ctx.sync();
                        })
                    }
                    document.body.style.setProperty(`--progress`, `360deg`);
                    spinner.remove();
                }, 5);
            }
        },
    });
    input.click();
}

/**
 * Update Word properties
 * @param style the Word.Style object that'll be updated
 * @param item the Word.Style object (usually, fetched from a JSON file) that'll be copied to the `style` property
 */
async function updateProperties(style: Word.Style, item: Word.Style) {
    for (const firstKey in item) {
        if (typeof item[firstKey as "font"] === "object") {
            // Nested property
            for (const secondKey in item[firstKey as "font"]) {
                if (firstKey === "borders" &&!isNaN(+secondKey)) {
                    // Borders object can be an array. In this case, we have to get all the properties in this array, and update them
                    for (const thirdKey in item.borders[secondKey as "items"]) {
                        try {
                            // @ts-ignore - I don't know what I was trying to tell Typescript there
                            if (style.borders.items[+(secondKey as "0")][thirdKey as "0"] !== item.borders[+secondKey as "items"][thirdKey as "0"]) style.borders.items[+(secondKey as "0")][thirdKey as "0"] = item.borders[+secondKey as "items"][thirdKey as "0"];
                        } catch (ex) {
                            console.warn(ex);
                        }
                    }
                } else {
                    try {
                        if (firstKey === "shading" && secondKey === "foregroundPatternColor" && item[firstKey as "font"][secondKey as "size"] === null) continue; // Fix to avoid a generic error in the Office.JS runtime
                        if (style[firstKey as "font"][secondKey as "size"] !==item[firstKey as "font"][secondKey as "size"]) style[firstKey as "font"][secondKey as "size"] =item[firstKey as "font"][secondKey as "size"]; // Update only if the value is differnt.
                    } catch (ex) {
                        console.warn(ex);
                    }
                }
            }
        } else {
            try {
                if (style[firstKey as "visibility"] !== item[firstKey as "visibility"]) style[firstKey as "visibility"] = item[firstKey as "visibility"];
            } catch (ex) {
                console.warn(ex);
            }
        }
    }
}