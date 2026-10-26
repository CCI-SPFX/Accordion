import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
// The SCSS module is processed by the SPFx build pipeline.
// @ts-ignore: TypeScript may not have the generated SCSS module declaration.
import styles from "./Accordion.module.scss";
import { IAccordionProps } from "./IAccordionProps";

interface IListItem {
  Id: number;
  Title: string;
  Description?: string;
  OrderNumber?: number;
}

const Accordion: React.FC<IAccordionProps> = (props) => {
  const [allItems, setAllItems] = useState<IListItem[]>([]);
  const [searchText, setSearchText] = useState<string>("");
  const [openId, setOpenId] = useState<number | null>(null);
  const [loading, setLoading] = useState<boolean>(true);







  // Detects whether a value contains HTML tags (rich text) or is plain text
const isHtml = (value: string): boolean => /<\/?[a-z][\s\S]*?>/i.test(value);

// Converts plain text to safe HTML (keeps line breaks), passes rich text through
const toHtml = (value?: string): string => {
  if (!value) return "";
  if (isHtml(value)) return value;
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\r?\n/g, "<br/>");
};

// Strips HTML so search only matches visible text, not tags
const toPlainText = (value?: string): string => {
  if (!value) return "";
  if (!isHtml(value)) return value;
  return new DOMParser().parseFromString(value, "text/html").body.textContent || "";
};


  // Fetch list items whenever the selected list changes
  useEffect(() => {
    if (!props.listName) {
      setAllItems([]);
      setLoading(false);
      return;
    }

    setLoading(true);
    setOpenId(null);

    const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      props.listName,
    )}')/items?$select=Id,Title,Description,OrderNumber&$orderby=OrderNumber asc&$top=5000`;

    props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: IListItem[] }) => {
        setAllItems(data.value || []);
        setLoading(false);
      })
      .catch(() => {
        setAllItems([]);
        setLoading(false);
      });
  }, [props.listName]);

  // Filter by search text (matched against the button/title text)
  // const filteredItems = allItems.filter(
  //   (it) =>
  //     (it.Title || "").toLowerCase().indexOf(searchText.toLowerCase()) !== -1,
  // );
  // Filter by search text (matched against the title OR the description)
const filteredItems = allItems.filter((it) => {
  const search = searchText.toLowerCase();
  return (
    // (it.Title || "").toLowerCase().indexOf(search) !== -1 ||  (it.Description || "").toLowerCase().indexOf(search) !== -1
    (it.Title || "").toLowerCase().indexOf(search) !== -1 || toPlainText(it.Description).toLowerCase().indexOf(search) !== -1
  );
});
  // Respect the "items to show" property pane setting
  const visibleItems = filteredItems.slice(0, props.itemsToShow || 4);

  const toggleItem = (id: number): void => {
    setOpenId((prev) => (prev === id ? null : id));
  };

  // Button radius: full radius when closed; top corners rounded / bottom corners sharp when open
  const getButtonStyle = (isOpen: boolean): React.CSSProperties => ({
    fontSize: `${props.buttonFontSize}px`,
    fontWeight: props.buttonFontWeight as React.CSSProperties["fontWeight"],
    backgroundColor: props.buttonBackgroundColor,
    color: props.buttonTextColor,
    borderRadius:
      isOpen ?
        `${props.buttonRadius}px ${props.buttonRadius}px 0px 0px`
      : `${props.buttonRadius}px`,
    minHeight: `${props.itemsHeight}px`,
  });

  // Panel: sharp top corners (matches the open button's bottom edge), rounded bottom corners
  const panelStyle: React.CSSProperties = {
    fontSize: `${props.panelFontSize}px`,
    fontWeight: props.panelFontWeight as React.CSSProperties["fontWeight"],
    backgroundColor: props.panelBackgroundColor,
    color: props.panelTextColor,
    borderRadius: `0px 0px ${props.buttonRadius}px ${props.buttonRadius}px`,
  };

  const searchStyle: React.CSSProperties = {
    fontSize: `${props.buttonFontSize}px`,
    fontWeight: props.buttonFontWeight as React.CSSProperties["fontWeight"],
    borderRadius: `${props.buttonRadius}px`,
    minHeight: `${props.itemsHeight}px`,
  };

  return (
    <div
      className={styles.accordionWrap}
      style={{ gap: `${props.itemsGap}px` }}
    >
      <div className={styles.searchWrap}>
        <div
          className={styles.searchInputContainer}
          style={{ display: `${props.searchBarEnabled ? "flex" : "none"}` }}
        >
          <svg
            className={styles.searchIcon}
            viewBox="0 0 24 24"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
            aria-hidden="true"
          >
            <path
              d="M15.5 14h-.79l-.28-.27a6.5 6.5 0 1 0-.7.7l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0A4.5 4.5 0 1 1 14 9.5 4.5 4.5 0 0 1 9.5 14z"
              fill="currentColor"
            />
          </svg>
          <input
            type="text"
            name="search"
            id="search"
            className={styles.searchInput}
            style={searchStyle}
            placeholder={props.placeholderText}
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
          />
        </div>
      </div>

      {loading && <div className={styles.notFound}>Loading...</div>}

      {!loading && !props.listName && (
        <div className={styles.notFound}>
          Please select a list from the web part properties.
        </div>
      )}

      {!loading && props.listName && visibleItems.length === 0 && (
        <div className={styles.notFound}>Not found</div>
      )}

      {!loading &&
        visibleItems.map((item) => {
          const isOpen = openId === item.Id;
          return (
            <div className={styles.accordionItem} key={item.Id}>
              <button
                type="button"
                className={styles.accordionButton}
                style={getButtonStyle(isOpen)}
                onClick={() => toggleItem(item.Id)}
                aria-expanded={isOpen}
              >
                <span className={styles.buttonText}>{item.Title}</span>
                <span
                  className={`${styles.chevron} ${isOpen ? styles.chevronOpen : ""}`}
                >
                  <svg
                    width="15"
                    height="15"
                    viewBox="0 0 12 12"
                    fill="none"
                    xmlns="http://www.w3.org/2000/svg"
                  >
                    <path
                      d="M2 4L6 8L10 4"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                </span>
              </button>

              {/* {isOpen && (
                <div className={styles.accordionPanel} style={panelStyle}>
                  {item.Description || "No additional details for this item."}
                </div>
              )} */}
              {isOpen && (
  <div
    className={styles.accordionPanel}
    style={panelStyle}
    dangerouslySetInnerHTML={{
      __html:
        toHtml(item.Description) ||
        "No additional details for this item.",
    }}
  />
)}
            </div>
          );
        })}
    </div>
  );
};

export default Accordion;
