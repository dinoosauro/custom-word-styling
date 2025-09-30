<script lang="ts">
    import { lang } from "../Scripts/Language";
    import Card from "./Card.svelte";
    import DeleteButton from "./DeleteButton.svelte";

    let { table }: { table: Word.Table } = $props();
    /**
     * The row of the cell that is being edited
     */
    let selectedRow = 1;
    /**
     * The column of the cell that is being edited.
     * Note that we have no way to know if a column in this position actually exists, since tables might not have the same number of columns in each row.
     */
    let selectedColumn = 1;
    /**
     * Properties of the border
     */
    let borderProps = {
        color: "#000000" as string | null,
        width: 1,
    };
    /**
     * An object that contains all the position of the padding that needs to be updated
     */
    let updatePadding = {
        Top: false,
        Left: false,
        Right: false,
        Bottom: false,
    };
    /**
     * The new padding for the single cell
     */
    let paddingWidth = 0;
    /**
     * Properties that are applied only to the single cell
     */
    let singleCellProps = {
        verticalAlignment: "Center",
        horizontalAlignment: "Centered",
        shadingColor: null,
    };
    /**
     * The border that is being edited.
     * Note that a single border can be selected, since the user can also choose "All" from the dropdown menu (and therefore putting checkboxes would be redundant).
     */
    let selectedBorder = "Top";
    /**
     * A Spinner element at the center of the page
     */
    const spinner = document.createElement("div");
    spinner.classList.add("spinner");
    /**
     * If true, the entire column will be updated in the "single cell" mode
     */
    let selectAllRows = false;
    /**
     * If true, the entire row will be updated in the "single cell" mode
     */
    let selectAllColumns = false;
    /**
     * Get all the single cells that should be edited
     * @param ctx the Word request context
     */
    async function getCells(ctx: Word.RequestContext) {
        const range = ctx.document.getSelection().load();
        await ctx.sync();
        const tables = range.tables.load();
        await ctx.sync();
        let output: Word.TableCell[] = [];
        for (const table of tables.items) {
            if (selectAllRows || selectAllColumns) { // We need to get all the cells in a specific row
                const rows = table.rows.load();
                await ctx.sync();
                if (selectAllColumns && !selectAllRows) { // We now need to get the cell at the selectedRow - 1 position in the array
                    if (typeof rows.items[selectedRow - 1] === "undefined") continue; // If the table doesn't have that row, continue
                    const cells = rows.items[selectedRow - 1].cells.load();
                    await ctx.sync();
                    output.push(...cells.items);
                    continue;
                }
                // Otherwise, we need to get a cell per row
                for (const row of rows.items) {
                    const cells = row.cells.load();
                    await ctx.sync();
                    if (selectAllRows && selectAllColumns) { // We can just add all the cells to the output array
                        output.push(...cells.items);
                    } else {
                        typeof cells.items[selectedColumn - 1] !==
                            "undefined" &&
                            output.push(cells.items[selectedColumn - 1]); // If the cell exists, let's add it (since each row can have a different amount of columns compared to the others)
                    }
                }
            } else
                output.push(table.getCell(selectedRow - 1, selectedColumn - 1)); // If both options are disabled, we just need to get the requested cell.
        }
        return output;
    }
</script>

<label class="flex hcenter gap">
    {lang("Table alignment")}:
    <div class="selectContainer">
        <select bind:value={table.alignment}>
            <option value="Left">{lang("Left")}</option>
            <option value="Centered">{lang("Centered")}</option>
            <option value="Right">{lang("Right")}</option>
        </select>
    </div>
</label><br />
<label class="flex hcenter gap">
    {lang("Vertical alignment")} ({lang("all cells")}):
    <div class="selectContainer">
        <select bind:value={table.verticalAlignment}>
            <option value="Top">{lang("Top")}</option>
            <option value="Center">{lang("Center")}</option>
            <option value="Bottom">{lang("Bottom")}</option>
        </select>
    </div>
</label><br />
<label class="flex hcenter gap">
    {lang("Horizontal alignment")} ({lang("all cells")}):
    <div class="selectContainer">
        <select bind:value={table.horizontalAlignment}>
            <option value="Left">{lang("Left")}</option>
            <option value="Centered">{lang("Centered")}</option>
            <option value="Right">{lang("Right")}</option>
            <option value="Justified">{lang("Justified")}</option>
        </select>
    </div>
</label><br />
<Card secondCard={true}>
    <h3>{lang("Cell-specific styling")}:</h3>
    <p>
        {lang(
            'Since you might want to edit different cells, you\'ll need to click on the "Update" button inside each card to apply the settings to only the selected cell',
        )}.
    </p>
    <label class="flex hcenter gap">
        {lang("Cell of row")}:
        <input
            type="number"
            bind:value={selectedRow}
            min="1"
            max={table.rowCount}
        />
        {lang("and column")}:
        <input type="number" bind:value={selectedColumn} min="1" />
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={selectAllColumns}>{lang("Select all the cells in the selected row")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={selectAllRows}>{lang("Select all the cells in the selected column")}
    </label><br>
    <Card>
        <h4>{lang("Border")}:</h4>
        <label class="flex hcenter gap">
            {lang("Edit the following border")}:
            <div class="selectContainer">
                <select bind:value={selectedBorder}>
                    <option value="Top">{lang("Top")}</option>
                    <option value="Left">{lang("Left")}</option>
                    <option value="Bottom">{lang("Bottom")}</option>
                    <option value="Right">{lang("Right")}</option>
                    <option value="InsideHorizontal"
                        >{lang("Inside Horizontal")}</option
                    >
                    <option value="InsideVertical"
                        >{lang("Inside Vertical")}</option
                    >
                    <option value="Inside">{lang("Inside")}</option>
                    <option value="Outside">{lang("Outside")}</option>
                    <option value="All">{lang("All")}</option>
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Color")}:
            <input type="color" bind:value={borderProps.color} />
            <DeleteButton
                callback={(e) => {
                    borderProps.color = null;
                    const possibleItem = (e.target as HTMLElement)
                        .closest("label")
                        ?.querySelector("input[type=color]");
                    if (possibleItem)
                        (possibleItem as HTMLInputElement).value = "#000000";
                }}
            ></DeleteButton>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Width")}:
            <input type="number" bind:value={borderProps.width} />
        </label><br />
        <button
            onclick={async () => {
                document.body.append(spinner);
                setTimeout(async () => {
                    await Word.run(async (ctx) => {
                        for (const cell of await getCells(ctx)) {
                            cell.load();
                            await ctx.sync();
                            const border = cell
                                .getBorder(selectedBorder as "Top")
                                .load();
                            await ctx.sync();
                            for (const prop in borderProps)
                                if (border[prop as "color"] !== null)
                                    border[prop as "color"] = borderProps[
                                        prop as "color"
                                    ] as string;
                            await ctx.sync();
                        }
                    });
                    spinner.remove();
                }, 1);
            }}>{lang("Update border")}</button
        >
    </Card><br />
    <Card>
        <h4>{lang("Cell padding")}:</h4>
        <label class="flex gap" style="flex-wrap: wrap;">
            {lang("Update the following paddings")}:
            {#each Object.keys(updatePadding) as text}
                <label class="flex hcenter gap" style="gap: 5px;">
                    <input
                        type="checkbox"
                        bind:checked={updatePadding[text as "Top"]}
                    />
                    {lang(text)}
                </label>
            {/each}
        </label><br />
        <label class="flex hcenter gap">
            {lang("Padding width (in points)")}:
            <input type="number" bind:value={paddingWidth} />
        </label><br />
        <button
            onclick={async () => {
                document.body.append(spinner);
                setTimeout(async () => {
                    await Word.run(async (ctx) => {
                        for (const cell of await getCells(ctx)) {
                            for (const prop in updatePadding)
                                updatePadding[prop as "Top"] &&
                                    cell.setCellPadding(
                                        prop as "Top",
                                        paddingWidth,
                                    );
                            await ctx.sync();
                        }
                    });
                    spinner.remove();
                }, 1);
            }}>{lang("Update padding")}</button
        >
    </Card><br />
    <Card>
        <h4>{lang("Other single cell styles")}:</h4>
        <label class="flex hcenter gap">
            {lang("Vertical alignment")} ({lang("single cell")}):
            <div class="selectContainer">
                <select bind:value={singleCellProps.verticalAlignment}>
                    <option value="Top">{lang("Top")}</option>
                    <option value="Center">{lang("Center")}</option>
                    <option value="Bottom">{lang("Bottom")}</option>
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Horizontal alignment")} ({lang("single cell")}):
            <div class="selectContainer">
                <select bind:value={singleCellProps.horizontalAlignment}>
                    <option value="Left">{lang("Left")}</option>
                    <option value="Centered">{lang("Centered")}</option>
                    <option value="Right">{lang("Right")}</option>
                    <option value="Justified">{lang("Justified")}</option>
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Cell color")}:
            <input type="color" bind:value={singleCellProps.shadingColor} />
            <DeleteButton
                callback={(e) => {
                    singleCellProps.shadingColor = null;
                    const possibleItem = (e.target as HTMLElement)
                        .closest("label")
                        ?.querySelector("input[type=color]");
                    if (possibleItem)
                        (possibleItem as HTMLInputElement).value = "#000000";
                }}
            ></DeleteButton>
        </label><br />
        <button
            onclick={async () => {
                document.body.append(spinner);
                setTimeout(async () => {
                    await Word.run(async (ctx) => {
                        for (const cell of await getCells(ctx)) {
                            for (const prop in singleCellProps) {
                                if (
                                    singleCellProps[prop as "shadingColor"] !==
                                    null
                                )
                                    cell[prop as "shadingColor"] =
                                        singleCellProps[
                                            prop as "shadingColor"
                                        ] as unknown as string;
                            }
                            await ctx.sync();
                        }
                    });
                    spinner.remove();
                }, 1);
            }}>{lang("Update cell styling")}</button
        >
    </Card>
</Card>
