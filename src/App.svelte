<script lang="ts">
  import { getContext, onMount } from "svelte";
  import AddStyleDialog from "./lib/AddStyleDialog.svelte";
  import FontChange from "./lib/FontChange.svelte";
  import ParagraphChange from "./lib/ParagraphChange.svelte";
  import BorderChange from "./lib/BorderChange.svelte";
  import ShadingChange from "./lib/ShadingChange.svelte";
  import TableChange from "./lib/TableChange.svelte";
  import ListOrderChange from "./lib/ListOrderChange.svelte";
  import { fade } from "svelte/transition";
  import { cubicInOut } from "svelte/easing";
  import ImportStyles from "./Scripts/ImportStyles";
  import ExportStyles from "./Scripts/ExportStyles";
    import NormalItemChange from "./lib/NormalItemChange.svelte";
    import type { HelperType } from "./Scripts/HelperType";
    import HelperDialogs from "./lib/HelperDialogs.svelte";
    import { lang, updateOfficeReady } from "./Scripts/Language";
    import OpenSourceDialog from "./lib/OpenSourceDialog.svelte";
    import ParagraphListEdit from "./lib/ParagraphListEdit.svelte";
    import Card from "./lib/Card.svelte";
    import AddShapeMainContent from "./lib/AddShapeMainContent.svelte";
    import EditTableCell from "./lib/EditTableCell.svelte";
    import EditSelectedShape from "./lib/EditSelectedShape.svelte";
  /**
   * The currently-selected style that is being edited.
   *
   * **Applies to:** `styleChange` section
   */
  let index = $state(-1);
  /**
   * The section of the add-in that is being displayed
   */
  let appSection:
    | "none"
    | "styleChange"
    | "paragraphChange"
    | "styleApply"
    | "styleExport"
    | "newShape" = $state("none");
  /**
   * The `items` property of the StyleCollection fetched from Word API.
   * This object is provided only so that the style name can be read, but must not be used, since updating a property there would also change it in Word, making the "Discard" function useless.
   */
  let referenceItems: Word.Style[] = [];
  /**
   * Deep clone of `referenceItems`, generated gradually while the user changes the selection.
   */
  let availableItems: Word.Style[] = [];
  /**
   * The custom styling tab that was chosen by the user.
   *
   * **Applies to:** `styleChange` section
   */
  let currentlyChosenTab:
    | "border"
    | "font"
    | "list"
    | "paragraph"
    | "shading"
    | "table"
    | "general" = $state("font");
  /**
   * If the dialog to add a new style should be visible
   */
  let showAddDialog = $state(false);
  /**
   * The string that is used for the select in the `styleChange` section. It should be changed after the `referenceItem` list has been changed, so that a re-render of the options inside the select can be triggered.
   */
  let rerenderSelect = $state(`SelectBlock-${Date.now()}`);
  /**
   * If changed, it'll re-render everything except the header.
   */
  let forceReRender = $state(`Entire-${Date.now()}`);
  /**
   * If the `List` tab of the currently-selected style should be one of the pickable tabs.
   *
   * **Applies to:** `styleChange` tab
   */
  let shouldListTabBeVisible = $state(false);
  /**
   * An array of bool values. If true, the style at that position should be exported.
   *
   * **Applies to:** `styleExport` tab
   */
  let propertiesToExport: boolean[] = [];
  /**
   * Possible options when importing a custom style
   *
   * **Applies to:** `styleApply` tab
   */
  let mergeStyleOption: "CreateNew" | "Overwrite" | "Ignore" = "Overwrite";
  /**
   * Information about the text selected by the user, used so that it can be customized in the `paragraphChange` section.
   *
   * **Applies to:** `paragraphChange` tab
   */
  let selectedParagraph: {
    font: Word.Font;
    paragraphs: Word.Paragraph[];
    lists: Word.List[],
    tables: Word.Table[],
    shapes: Word.Shape[]
  };
  /**
   * The style that should be applied to the selected text.
   *
   * **Applies to:** `styleApply` tab
   */
  let selectedChangeStyle = "0";
  /**
   * If the dialog to go back to the home should be triggered.
   */
  let showHomeDialog = $state(false);
  /**
   * The link that needs to be displayed in the "Download file" dialog.
   * It's false if the dialog doesn't need to be displayed, either because the user hasn't asked to download a file or because the OpenBrowserWindowApi is supported.
   */
  let downloadLink: false | string = $state(false);
  /**
   * If not undefined, a dialog with some useful information will be shown.
   */
  let helperProp: HelperType | undefined = $state(undefined);
  /**
   * Update the `availableItems` property with the clone of the style selected by the user
   * @param nextIndex the position in the `referenceItems` object of the item to clone
   * @param ctx the Context used to get the content inside `referenceItems`
   */
  async function getIndexReady(nextIndex: number, ctx: Word.RequestContext) {
    // @ts-ignore
    referenceItems[nextIndex].load({$all: true, borders: true, font: true, listTemplate: true, paragraphFormat: true, shading: true, tableStyle: true});
    await ctx.sync();
    referenceItems[nextIndex].listTemplate.listLevels.load();
    await ctx.sync();
    let items = referenceItems[nextIndex].listTemplate.listLevels.items;
    if (items) {
      // We'll deep clone also its font, and we'll add it manually in the JSON file since by default they're discarded
      let prevFont: Word.Font[] = items.map((i) => i?.font);
      items = JSON.parse(JSON.stringify(items));
      for (let i = 0; i < items.length; i++) {
        // @ts-ignore
        if (prevFont[i]) items[i].font = JSON.parse(JSON.stringify(prevFont[i]));
      }
    }
    // Now we'll deep clone the object
    availableItems[nextIndex] = JSON.parse(
      JSON.stringify(referenceItems[nextIndex]),
    );
    // @ts-ignore
    if (items !== undefined) availableItems[nextIndex].listTemplate.listLevels = { items };
    // @ts-ignore
    if (referenceItems[nextIndex].borders.items?.length > 0) availableItems[nextIndex].borders = {
      items: JSON.parse(JSON.stringify(referenceItems[nextIndex].borders.items)),
      insideBorderColor: referenceItems[nextIndex].borders.insideBorderColor,
      insideBorderWidth: referenceItems[nextIndex].borders.insideBorderWidth,
      insideBorderType: referenceItems[nextIndex].borders.insideBorderType,
      outsideBorderColor: referenceItems[nextIndex].borders.outsideBorderColor,
      outsideBorderWidth: referenceItems[nextIndex].borders.outsideBorderWidth,
      outsideBorderType: referenceItems[nextIndex].borders.outsideBorderType,
    }
    shouldListTabBeVisible =
      Array.isArray(availableItems[nextIndex].listTemplate.listLevels.items) &&
      availableItems[nextIndex].listTemplate.listLevels.items.length > 0; // Update the "list" tab visibility by checking that there's at least one list item that can be edited
    index = nextIndex; // And **finally** update the index.
  }
  /**
   * Get all the styles
   * @param suggestedIndex set the style at this index as the selected one
   */
  async function loadStyle(suggestedIndex = 0) {
    await Word.run(async (ctx) => {
      const styles = ctx.document.getStyles().load();
      await ctx.sync();
      referenceItems = styles.items;
      suggestedIndex !== -1 && (await getIndexReady(suggestedIndex, ctx));
    });
  }

  // Dark/light mode change part

  /**
   * CSS properties that should be changed from light to dark mode
   */
  const itemsToChange = [
    "text",
    "background",
    "card",
    "input",
    "accent",
    "hover-filter",
    "active-filter",
  ];
  const darkThemeVariant = itemsToChange.map((i) =>
    getComputedStyle(document.body).getPropertyValue(`--${i}`),
  );
  const lightThemeVariant = [
    "#151515",
    "#fafafa",
    "#d2d2d2",
    "#b4b4b4",
    "#bbc698",
    "85%",
    "75%",
  ];
  let isLightTheme = false;
  /**
   * Switch between the light and the dark theme
   * @param skipLocalStorageSaving if true, the preference won't be saved in the LocalStorage. This should be done in the case the preference is fetched from the Word UI, and not from user input.
   */
  function changeTheme(skipLocalStorageSaving?: boolean) {
    isLightTheme = !isLightTheme;
    titleImage.src = isLightTheme ? "./logo_dark.svg" : "./logo_light.svg";
    for (let i = 0; i < itemsToChange.length; i++) {
      document.body.style.setProperty(
        `--${itemsToChange[i]}`,
        isLightTheme ? lightThemeVariant[i] : darkThemeVariant[i],
      );
    }
    !skipLocalStorageSaving && localStorage.setItem(
      "CustomWordStyle-Theme",
      isLightTheme ? "light" : "dark",
    );
  }
  onMount(() => {
    console.error = (err) => (errorDialog = err.toString());
    window.addEventListener("error", (e) => {
      errorDialog = `${e.error.toString()}\n\nFrom line ${e.lineno}, column ${e.colno} of ${e.filename}`;
      document.querySelector(".spinner")?.remove();
    });
    window.addEventListener("unhandledrejection", (e) => {
      errorDialog = e.reason;
      document.querySelector(".spinner")?.remove();
    })
    Office.onReady().then(() => {
      updateOfficeReady();
      if (
        (!Office.context.officeTheme.isDarkTheme &&
          localStorage.getItem("CustomWordStyle-Theme") !== "dark") ||
        localStorage.getItem("CustomWordStyle-Theme") === "light"
      ) {
        changeTheme(true);
      }
    });
    setTimeout(() => (forceReRender = `Entire-${Date.now()}`), 150);
    
  });
  /**
   * A Spinner element at the center of the page
   */
  const spinner = document.createElement("div");
  spinner.classList.add("spinner");
  /**
   * Update the `selectParagraph` property by getting the text the user has selected, and update the app section
   */
  async function getSelection() {
    await Word.run(async (ctx) => {
      let paragraph = ctx.document.getSelection().load();
      await ctx.sync();
      paragraph.font.load();
      await ctx.sync();
      paragraph.paragraphs.load();
      await ctx.sync();
      paragraph.lists.load();
      await ctx.sync();
      paragraph.tables.load();
      await ctx.sync();
      // @ts-ignore
      paragraph.shapes.load({$all: true, textFrame: true, textWrap: true, fill: true});
      await ctx.sync();
      selectedParagraph = {
        paragraphs: JSON.parse(JSON.stringify(paragraph.paragraphs.items)),
        font: JSON.parse(JSON.stringify(paragraph.font)),
        lists: JSON.parse(JSON.stringify(paragraph.lists.items)),
        tables: JSON.parse(JSON.stringify(paragraph.tables.items)),
        shapes: JSON.parse(JSON.stringify(paragraph.shapes.items))
      };
      appSection = "paragraphChange";
    });
  }
  /**
   * The icon of the website, dynamically updated when the user switches from light to dark mode and viceversa
   */
  let titleImage: HTMLImageElement;
  /**
   * If true, the dialog with the licenses will be shown
   */
  let showLicenseDialog = $state(false);
  let errorDialog: false | string = $state(false);
</script>


{#key forceReRender}
  <header>
    <div class="flex hcenter gap">
      <img
        bind:this={titleImage}
        onclick={() => {
          showHomeDialog = true;
        }}
        class="hover"
        style="width: 48px; height: 48px"
        src="./logo_light.svg"
        alt="Website icon. Click on it to go back to the Selection tab"
      />
      <h1>Custom Word Styling</h1>
    </div>
    <p>{lang("Change Word styles using the Office.JS API")}</p>
  </header>

  <div>
    {#if appSection === "none"}
      <Card>
        <h2>{lang("Change styles")}:</h2>
        <p>{lang("Change the styles of your document, and create new ones.")}</p>
        <button
          onclick={async () => {
            document.body.append(spinner);
            setTimeout(async () => {
              await loadStyle();
              appSection = "styleChange";
              spinner.remove();
            }, 1);
          }}>{lang("Load styles")}</button
        >
      </Card>
      <br />
      <Card>
        <h2>{lang("Apply style")}:</h2>
        <p>{lang("Apply a style to the text you've selected on Word")}.</p>
        <button
          onclick={async () => {
            document.body.append(spinner);
            setTimeout(async () => {
              await loadStyle(-1);
              appSection = "styleApply";
              spinner.remove();
            }, 1);
          }}>{lang("Load styles")}</button
        >
      </Card>
      <br />
      <Card>
        <h2>{lang("Change selection")}:</h2>
        <p>{lang("Change how the selected text looks, without creating a new style")}.</p>
        <button
          onclick={async () => {
            document.body.append(spinner);
            setTimeout(async () => {
              await getSelection();
              spinner.remove();
            }, 1);
          }}>{lang("Get selection")}</button
        >
      </Card><br>
      <Card>
        <h2>{lang("Customize shapes")}:</h2>
        <p>{lang("Create a new shape with a custom border radius, background color (both plain and gradient) or background image")}.</p>
        <button onclick={() => (appSection = "newShape")}>{lang("Create new shape")}</button>
      </Card>
      <br />
      <Card>
        <h2>{lang("Export styles")}:</h2>
        <p>{lang("You'll be able to choose which styles to export in a JSON file")}.</p>
        <button
          onclick={async () => {
            document.body.append(spinner);
            setTimeout(async () => {
              await loadStyle(0);
              appSection = "styleExport";
              spinner.remove();
            }, 1);
          }}>{lang("Export styles")}</button
        >
      </Card>
      <br />
      <Card>
        <h2>{lang("Import styles")}</h2>
        <p>
          {lang("If you've exported a JSON file before, you can import its styles here")}.
        </p>
        <label class="flex hcenter gap">
          {lang("If a style with the same name is found")},
          <div class="selectContainer">
            <select bind:value={mergeStyleOption}>
              <option value="Overwrite">{lang("overwrite it")}</option>
              <option value="Ignore">{lang("ignore it")}</option>
              <option value="CreateNew">{lang("keep both of them")}</option>
            </select>
          </div>
        </label><br />
        <button
          onclick={() => ImportStyles(mergeStyleOption)}>{lang("Pick JSON file")}</button
        >
     </Card>
    {:else if appSection === "styleChange"}
      <div class="flex hcenter gap">
        {#key rerenderSelect}
          <div class="selectContainer">
            <select
              onchange={async (e) => {
                document.body.append(spinner);
                setTimeout(async () => {
                  await loadStyle(+(e.target as HTMLSelectElement).value);
                  spinner.remove();
                }, 0);
              }}
            >
              {#each referenceItems as item, i}
                <option value={i}>{item.nameLocal}</option>
              {/each}
            </select>
          </div>
        {/key}
        <button
          style="width: fit-content"
          onclick={() => {
            showAddDialog = true;
          }}>+</button
        >
      </div>
      <br />
      <Card>
        <div class="flex gap" style="overflow: auto">
          {#each [["font", "Font options"], ["paragraph", "Paragraph options"], ["border", "Border options"], ["general", "General options"], ["shading", "Shading options"], ...(availableItems[index].type === "Table" ? [["table", "Table options"]] : []), ...(shouldListTabBeVisible ? [["list", "List options"]] : [])] as [key, title]}
            <button
              class="card secondCard chip"
              style={currentlyChosenTab === key
                ? "background-color: var(--accent)"
                : "background-color: var(--input)"}
              onclick={() => (currentlyChosenTab = key as "list")}
            >
              {lang(title)}
            </button>
          {/each}
        </div>
      </Card>
      <br />
      {#if currentlyChosenTab === "font"}
        <Card>
          <h2>{lang("Font")}:</h2>
          <FontChange sourceFont={availableItems[index].font}></FontChange>
        </Card>
      {/if}
      {#if currentlyChosenTab === "paragraph"}
        <Card>
          <h2>{lang("Paragraph")}:</h2>
          <ParagraphChange
            isParagraphFormat={true}
            sourceParagraph={availableItems[index].paragraphFormat}
          ></ParagraphChange>
        </Card>
      {/if}
      {#if currentlyChosenTab === "border"}
        <Card>
          <h2>{lang("Border")}:</h2>
          <BorderChange sourceBorder={availableItems[index].borders}
          ></BorderChange>
        </Card>
      {/if}
      {#if currentlyChosenTab === "shading"}
        <Card>
          <h2><span class="help" onclick={() => (helperProp = "Shading")}>{lang("Shading (background color)")}:</span></h2>
          <ShadingChange sourceShading={availableItems[index].shading}
          ></ShadingChange>
        </Card>
      {/if}
      {#if currentlyChosenTab === "table"}
        <Card>
          <h2>{lang("Table")}:</h2>
          <TableChange sourceTable={availableItems[index].tableStyle}
          ></TableChange>
        </Card>
      {/if}
      {#if shouldListTabBeVisible && currentlyChosenTab === "list"}
        <Card>
          <h2>{lang("List order")}:</h2>
          <ListOrderChange sourceList={availableItems[index].listTemplate}
          ></ListOrderChange>
        </Card>
      {/if}
      {#if currentlyChosenTab === "general"}
        <Card>
          <h2>{lang("General options")}:</h2>
          <NormalItemChange generalStyle={availableItems[index]}></NormalItemChange>
        </Card>
      {/if}
      <br /><br />
      <button
        onclick={async () => {
          document.body.append(spinner);
          setTimeout(async () => {
            await Word.run(async (ctx) => {
              document.body.append(spinner);
              const styles = ctx.document.getStyles().load();
              await ctx.sync();
              // Update the entries in the real Style object.
              for (const entry of [
                "borders",
                "font",
                "paragraphFormat",
                "shading",
                "tableStyle",
              ]) {
                const obj = styles.items[index][entry as "font"].load();
                await ctx.sync();
                for (const key in availableItems[index][entry as "font"]) {
                  if (Array.isArray(availableItems[index][entry as "borders"][key as "items"])) {
                    for (let i = 0; i < availableItems[index][entry as "borders"][key as "items"].length; i++) {
                      for (const secondKey in availableItems[index][entry as "borders"][key as "items"][i]) {
                        try {
                          if ((obj as unknown as Word.BorderCollection)[key as "items"][i][secondKey as "visible"] !== availableItems[index][entry as "borders"][key as "items"][i][secondKey as "visible"]) (obj as unknown as Word.BorderCollection)[key as "items"][i][secondKey as "visible"] = availableItems[index][entry as "borders"][key as "items"][i][secondKey as "visible"];
                        } catch (ex) {
                          console.warn(ex);
                        }
                      }
                    }
                  } else {
                    try {
                     if (obj[key as "size"] !== availableItems[index][entry as "font"][key as "size"]) obj[key as "size"] = availableItems[index][entry as "font"][key as "size"];
                    } catch (ex) {
                      console.warn(ex);
                    }
                  }
                }
                await ctx.sync();
              }
              if (shouldListTabBeVisible) {
                // Update also the content for each list level
                styles.items[index].listTemplate.listLevels.load();
                await ctx.sync();
                for (
                  let i = 0;
                  i <
                  availableItems[index].listTemplate.listLevels.items.length;
                  i++
                ) {
                  for (const key in availableItems[index].listTemplate
                    .listLevels.items[i]) {
                    if (key === "font") {
                      for (const fontKey in availableItems[index].listTemplate
                        .listLevels.items[i].font) {
                        styles.items[index].listTemplate.listLevels.items[
                          i
                        ].font[key as "size"] =
                          availableItems[index].listTemplate.listLevels.items[
                            i
                          ].font[fontKey as "size"];
                          await ctx.sync();
                      }
                    } else {
                      if (styles.items[index].listTemplate.listLevels.items[i][
                        key as "numberFormat"
                      ] !==
                        availableItems[index].listTemplate.listLevels.items[i][
                          key as "numberFormat"
                        ]) styles.items[index].listTemplate.listLevels.items[i][
                        key as "numberFormat"
                      ] =
                        availableItems[index].listTemplate.listLevels.items[i][
                          key as "numberFormat"
                        ];
                        await ctx.sync();
                    }
                  }
                  await ctx.sync();
                }
              }
              // Let's update also the basic properties of the style
              for (const property of ["baseStyle","nextParagraphStyle","priority","quickStyle","unhideWhenUsed","visibility"]) {
                if (styles.items[index][property as "baseStyle"] !== availableItems[index][property as "baseStyle"]) styles.items[index][property as "baseStyle"] = availableItems[index][property as "baseStyle"];
                  await ctx.sync();
              }
              await ctx.sync();
              spinner.remove();
            });
          }, 1);
        }}>{lang("Save")}</button
      ><br /><br />
      <u
        class="discard"
        onclick={async () => {
          document.body.append(spinner);
          setTimeout(async () => {
            // To discard it, we'll just load again the values
            await loadStyle(Math.max(0, index));
            forceReRender = `Entire-${Date.now()}`;
            spinner.remove();
          }, 1);
        }}>{lang("Discard edits")}</u
      >
    {:else if appSection === "styleApply"}
      <label class="flex hcenter gap">
        {lang("Apply this style")}:
        <div class="selectContainer">
          <select bind:value={selectedChangeStyle}>
            {#each referenceItems as item, i}
              <option value={i}>{item.nameLocal}</option>
            {/each}
          </select>
        </div>
      </label><br />
      <button
        onclick={async () => {
          document.body.append(spinner);
          setTimeout(async () => {
            await Word.run(async (ctx) => {
              const selection = ctx.document.getSelection().load();
              await ctx.sync();
              selection.style = referenceItems[+selectedChangeStyle].nameLocal;
              await ctx.sync();
              spinner.remove();
            });
          }, 1);
        }}>{lang("Apply")}</button
      >
    {:else if appSection === "paragraphChange"}
      <Card>
        <h2>{lang("Font")}:</h2>
        <FontChange sourceFont={selectedParagraph.font}></FontChange>
      </Card>
      <br />
      {#if selectedParagraph.paragraphs?.length > 0}
        <Card>
          <h2>{lang("Paragraph")}:</h2>
          <ParagraphChange
            isParagraphFormat={false}
            sourceParagraph={selectedParagraph.paragraphs[0]}
          ></ParagraphChange>
        </Card>
      {/if}
      {#if selectedParagraph.lists?.length > 0}
      <br>
      <Card>
        <h2>{lang("Lists")}:</h2>
        <ParagraphListEdit></ParagraphListEdit>
      </Card>
      {/if}
      {#if selectedParagraph.tables?.length > 0}
      <br>
      <Card>
        <h2>{lang("Table")}:</h2>
        <EditTableCell table={selectedParagraph.tables[0]}></EditTableCell>
      </Card>
      {/if}
      {#if selectedParagraph.shapes?.length > 0}
      <Card>
        <h2>{lang("Shapes")}:</h2>
        <EditSelectedShape shape={selectedParagraph.shapes[0]}></EditSelectedShape>
      </Card>
      {/if}
      <br />
      <button
        onclick={async () => {
          document.body.append(spinner);
          setTimeout(async () => {
            await Word.run(async (ctx) => {
              // Get the select range, and change the font and paragraph properties
              const range = ctx.document.getSelection().load();
              await ctx.sync();
              const font = range.font.load();
              await ctx.sync();
              for (const key in selectedParagraph.font) {
                try {
                  font[key as "size"] = selectedParagraph.font[key as "size"];
                  await ctx.sync();
                } catch (ex) {
                  console.warn(ex);
                }
              }
              await ctx.sync();
              if (selectedParagraph.paragraphs.length > 0) {
                const paragraphs = range.paragraphs.load();
                await ctx.sync();
                if (paragraphs.items?.length > 0) {
                  for (const paragraph of paragraphs.items) {
                    // Change all the paragraphs
                    for (const property of [
                      "alignment",
                      "firstLineIndent",
                      "outlineLevel",
                      "leftIndent",
                      "rightIndent",
                      "lineSpacing",
                      "spaceBefore",
                      "spaceAfter",
                    ]) {
                      paragraph[property as "spaceBefore"] =
                        selectedParagraph.paragraphs[0][
                          property as "spaceBefore"
                        ];
                    }
                    await ctx.sync();
                  }
                }
              }
              // Update tables
              if (selectedParagraph.tables?.length > 0) {
                const tables = range.tables.load();
                await ctx.sync();
                for (const table of tables.items) {
                  for (const prop of ["alignment", "verticalAlignment", "horizontalAlignment"]) table[prop as "alignment"] = selectedParagraph.tables[0][prop as "alignment"];
                }
                await ctx.sync();
              }
              // Update shapes
              if (selectedParagraph.shapes?.length > 0) {
                // @ts-ignore
                const shapes = range.shapes.load({$all: true, textFrame: true, textWrap: true, fill: true});
                await ctx.sync();
                for (const shape of shapes.items) {
                  shape.lockAspectRatio = false;
                  for (const prop in selectedParagraph.shapes[0]) {
                    if (typeof shape[prop as "fill"] === "object") { // Nested object to update
                      for (const secondProp in selectedParagraph.shapes[0][prop as "fill"]) {
                        if (secondProp === "textWrap" || secondProp === "hasText") continue; // Fix NotAllowed error
                        if (shape[prop as "fill"][secondProp as "transparency"] !== selectedParagraph.shapes[0][prop as "fill"][secondProp as "transparency"]) shape[prop as "fill"][secondProp as "transparency"] = selectedParagraph.shapes[0][prop as "fill"][secondProp as "transparency"];
                      await ctx.sync();
                      }
                    } else {
                      if (prop.startsWith("relative") || prop.endsWith("Relative")) continue;
                      if (shape[prop as "height"] !== selectedParagraph.shapes[0][prop as "height"]) {
                        shape[prop as "height"] = selectedParagraph.shapes[0][prop as "height"];
                      await ctx.sync();
                      }
                    }
                  }
                  await ctx.sync();
                }
              }
              spinner.remove();
            });
          }, 1);
        }}>{lang("Apply changes")}</button
      ><br /><br />
      <u
        class="discard"
        onclick={async () => {
          document.body.append(spinner);
          setTimeout(async () => {
            await getSelection();
            forceReRender = `Entire-${Date.now()}`;
            spinner.remove();
          }, 1);
        }}>{lang("Discard edits")}</u
      >
    {:else if appSection === "styleExport"}
      <p>
        {lang("Pick the properties you want to export. To avoid freezes when importing, choose only the properties you've changed.")}
      </p>
      <div>
      {#each referenceItems as item, i}
        <label class="flex hcenter gap">
          <input type="checkbox" bind:checked={propertiesToExport[i]} />
          {item.nameLocal}
        </label><br />
      {/each}
      </div><br>
      <div class="flex gap">
      <button
        onclick={async () => ExportStyles({propertiesToExport, urlCallback: (url) => {downloadLink = url}})}>{lang("Export")}</button
      >
      <button style="background-color: var(--input);" onclick={(e) => {
        // We'll untick everything if all the checkboxes are checked, otherwise we'll tick everything.
        let isEverythingChecked = true;
        for (let i = 0; i < referenceItems.length; i++) {
          if (!propertiesToExport[i]) {
            isEverythingChecked = false;
            break;
          }
        }
        for (let i = 0; i < referenceItems.length; i++) {
          propertiesToExport[i] = !isEverythingChecked;
        }
        // We'll now manually update the DOM with the new selection.
        const container = (e.target as HTMLButtonElement).closest("div")?.previousElementSibling?.previousElementSibling;
        if (container) {
          for (const checkbox of container.querySelectorAll("input[type=checkbox]")) {
            (checkbox as HTMLInputElement).checked = !isEverythingChecked;
          }
        }
      }}>{lang("Toggle all checkboxes")}</button>
    </div>
    {:else if appSection === "newShape"}
      <Card>
        <AddShapeMainContent></AddShapeMainContent>
      </Card>
    {/if}<br /><br />
    <div class="flex gap" style="flex-wrap: wrap;">
      <u
       onclick={() => changeTheme()}>{lang("Change theme")}</u
     >
      <u onclick={() => (showLicenseDialog = true)}>{lang("View open source licenses")}</u>
    </div>
    <p>{lang("Word and Office are trademarks of Microsoft. This project is no way affiliated or endorsed by Microsoft")}.</p>

    {#if showAddDialog}
      <AddStyleDialog
        callback={async (name, type, showAsQuickStyle) => {
          document.body.append(spinner);
          setTimeout(async () => {
            await Word.run(async (ctx) => {
              const style = ctx.document.addStyle(name, type as "Paragraph");
              style.visibility = true;
              style.unhideWhenUsed = true;
              style.quickStyle = showAsQuickStyle;
              showAddDialog = false;
            });
            await new Promise((res) => setTimeout(res, 150));
            await Word.run(async (ctx) => {
              await loadStyle();
              rerenderSelect = `SelectBlock-${Date.now()}`;
            });
            spinner.remove();
          }, 1);
        }}
      ></AddStyleDialog>
    {/if}
    {#if showHomeDialog}
      <div
        class="dialog"
        in:fade={{ duration: 400, easing: cubicInOut }}
        out:fade={{ duration: 400, easing: cubicInOut }}
      >
        <div>
          <h2>{lang("Do you want to go back to the home?")}</h2>
          <p>{lang("You'll lose all the unsaved changes.")}</p>
          <div class="flex gap">
            <button
              onclick={() => {
                showHomeDialog = false;
                appSection = "none";
              }}>{lang("Yes")}</button
            >
            <button
              style="background-color: var(--input);"
              onclick={() => (showHomeDialog = false)}>{lang("No")}</button
            >
          </div>
        </div>
      </div>
    {/if}
    {#if showLicenseDialog}
    <OpenSourceDialog callback={() => (showLicenseDialog = false)}></OpenSourceDialog>
    {/if}
    {#if downloadLink}
          <div
        class="dialog"
        in:fade={{ duration: 400, easing: cubicInOut }}
        out:fade={{ duration: 400, easing: cubicInOut }}
      >
        <div>
          <h2>{lang("Download file")}</h2>
          <p>{lang("Unfortunately, the platform you're using doesn't support file downloading. Copy the link displayed below, and then open it in your browser. Don't worry, your ultra-personal styles will always be private, and they aren't uploaded anywhere (all the download process happens locally).")}</p><br>
          <div class="secondCard" style="overflow: scroll;">
            <p style="white-space: pre;">{downloadLink}</p>
          </div><br>
          <div class="flex gap">
            <button onclick={() => navigator.clipboard.writeText(downloadLink as string)}>{lang("Copy")}</button>
            <button style="background-color: var(--input);" onclick={() => (downloadLink = false)}>{lang("Close")}</button>
          </div>
        </div>
      </div>
    {/if}
  </div>
  {#if helperProp}
  <HelperDialogs helperType={helperProp} callback={() => (helperProp = undefined)}></HelperDialogs>
  {/if}
  {#if errorDialog}
            <div
        class="dialog"
        in:fade={{ duration: 400, easing: cubicInOut }}
        out:fade={{ duration: 400, easing: cubicInOut }}
      >
        <div>
          <h2>{lang("An error occurred")} :(</h2>
          <p>{errorDialog}</p><br>
          <button onclick={() => (errorDialog = false)}>{lang("Close")}</button>
        </div>
        </div>
  {/if}
  {#if showLicenseDialog}
    <OpenSourceDialog callback={() => (showLicenseDialog = false)}></OpenSourceDialog>
  {/if}
{/key}
