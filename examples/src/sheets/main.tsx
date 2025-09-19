/**
 * Copyright 2023-present DreamNum Co., Ltd.
 * Licensed under the Apache License, Version 2.0
 */

import { createComponent } from "@lit/react";
import {
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
  UserManagerService,
} from "@univerjs/core";
import { FUniver } from "@univerjs/core/facade";
import { UniverDebuggerPlugin } from "@univerjs/debugger";
import { UniverDocsPlugin } from "@univerjs/docs";
import { UniverDocsUIPlugin } from "@univerjs/docs-ui";
import { UniverFormulaEnginePlugin } from "@univerjs/engine-formula";
import { UniverRenderEnginePlugin } from "@univerjs/engine-render";
import { DEFAULT_WORKBOOK_DATA_DEMO } from "@univerjs/mockdata";

import caES from "@univerjs/mockdata/locales/ca-ES";
import enUS from "@univerjs/mockdata/locales/en-US";
import esES from "@univerjs/mockdata/locales/es-ES";
import faIR from "@univerjs/mockdata/locales/fa-IR";
import frFR from "@univerjs/mockdata/locales/fr-FR";
import koKR from "@univerjs/mockdata/locales/ko-KR";
import ruRU from "@univerjs/mockdata/locales/ru-RU";
import viVN from "@univerjs/mockdata/locales/vi-VN";
import zhCN from "@univerjs/mockdata/locales/zh-CN";
import zhTW from "@univerjs/mockdata/locales/zh-TW";

import { UniverNetworkPlugin } from "@univerjs/network";
import { UniverRPCMainThreadPlugin } from "@univerjs/rpc";
import { UniverSheetsPlugin } from "@univerjs/sheets";
import { UniverSheetsConditionalFormattingPlugin } from "@univerjs/sheets-conditional-formatting";
import { UniverSheetsDataValidationPlugin } from "@univerjs/sheets-data-validation";
import { UniverSheetsFilterPlugin } from "@univerjs/sheets-filter";
import { UniverSheetsFormulaPlugin } from "@univerjs/sheets-formula";
import { UniverSheetsHyperLinkPlugin } from "@univerjs/sheets-hyper-link";
import { UniverSheetsNotePlugin } from "@univerjs/sheets-note";
import { UniverSheetsNumfmtPlugin } from "@univerjs/sheets-numfmt";
import { UniverSheetsSortPlugin } from "@univerjs/sheets-sort";
import { UniverSheetsTablePlugin } from "@univerjs/sheets-table";
import { UniverSheetsThreadCommentPlugin } from "@univerjs/sheets-thread-comment";
import { UniverSheetsUIPlugin } from "@univerjs/sheets-ui";
import { UniverSheetsZenEditorPlugin } from "@univerjs/sheets-zen-editor";
import { UniverUIPlugin } from "@univerjs/ui";
import { UniverVue3AdapterPlugin } from "@univerjs/ui-adapter-vue3";
import { UniverWebComponentAdapterPlugin } from "@univerjs/ui-adapter-web-component";

import { html, LitElement } from "lit";
import { customElement } from "lit/decorators.js";
import React from "react";

import { customRegisterEvent } from "./custom/custom-register-event";
import { UniverSheetsCustomShortcutPlugin } from "./custom/custom-shortcut";
import ImportCSVButtonPlugin from "./custom/import-csv-button";
import { ChatPlugin } from "./custom/chat-plugin";

import "../global.css";
import { render } from "@univerjs/design";

/* eslint-disable-next-line node/prefer-global/process */
const IS_E2E: boolean = !!process.env.IS_E2E;

const LOAD_LAZY_PLUGINS_TIMEOUT = 50;
const LOAD_VERY_LAZY_PLUGINS_TIMEOUT = 100;

export const mockUser = {
  userID: "Owner_qxVnhPbQ",
  name: "Owner",
  avatar:
    "data:image/png;base64,iVBORw0K... (rút gọn base64)...",
  anonymous: false,
  canBindAnonymous: false,
};

@customElement("my-univer-advanced")
class MyUniverAdvanced extends LitElement {
  override firstUpdated() {
    const container = this.renderRoot.querySelector(
      "#containerId"
    ) as HTMLDivElement;

    const univer = new Univer({
      darkMode: localStorage.getItem("local.darkMode") === "dark",
      locale: LocaleType.ZH_CN,
      locales: {
        [LocaleType.CA_ES]: caES,
        [LocaleType.EN_US]: enUS,
        [LocaleType.ES_ES]: esES,
        [LocaleType.FA_IR]: faIR,
        [LocaleType.FR_FR]: frFR,
        [LocaleType.KO_KR]: koKR,
        [LocaleType.RU_RU]: ruRU,
        [LocaleType.VI_VN]: viVN,
        [LocaleType.ZH_CN]: zhCN,
        [LocaleType.ZH_TW]: zhTW,
      },
      logLevel: LogLevel.VERBOSE,
    });

    const worker = new Worker(new URL("./worker.js", import.meta.url), {
      type: "module",
    });

    univer.registerPlugins([
      [UniverRPCMainThreadPlugin, { workerURL: worker }],
      [UniverDocsPlugin],
      [UniverRenderEnginePlugin],
      [UniverUIPlugin, { container }],
      [UniverWebComponentAdapterPlugin],
      [UniverVue3AdapterPlugin],
      [UniverDocsUIPlugin],
      [UniverSheetsPlugin, { notExecuteFormula: true, autoHeightForMergedCells: true }],
      [UniverSheetsUIPlugin],
      [UniverSheetsNumfmtPlugin],
      [UniverSheetsZenEditorPlugin],
      [UniverFormulaEnginePlugin, { notExecuteFormula: true }],
      [UniverSheetsFormulaPlugin, { notExecuteFormula: true }],
      [UniverSheetsDataValidationPlugin],
      [UniverSheetsConditionalFormattingPlugin],
      [UniverSheetsFilterPlugin],
      [UniverSheetsSortPlugin],
      [UniverSheetsHyperLinkPlugin],
      [UniverSheetsThreadCommentPlugin],
      [UniverSheetsTablePlugin],
      [UniverNetworkPlugin],
      [UniverSheetsNotePlugin],
      [ImportCSVButtonPlugin],
      [UniverSheetsCustomShortcutPlugin],
      [ChatPlugin],
    ]);

    console.log('Plugins registered, ChatPlugin should be loaded');

    if (IS_E2E) {
      univer.registerPlugin(UniverDebuggerPlugin, {
        fab: false,
        performanceMonitor: { enabled: false },
      });
    }

    const injector = univer.__getInjector();
    const userManagerService = injector.get(UserManagerService);
    userManagerService.setCurrentUser(mockUser);

    if (!IS_E2E) {
      univer.createUnit(
        UniverInstanceType.UNIVER_SHEET,
        DEFAULT_WORKBOOK_DATA_DEMO
      );
    }

    setTimeout(() => {
      import("./lazy").then((lazy) => {
        univer.registerPlugins(lazy.default());
      });
    }, LOAD_LAZY_PLUGINS_TIMEOUT);

    setTimeout(() => {
      import("./very-lazy").then((lazy) => {
        univer.registerPlugins(lazy.default());
      });
    }, LOAD_VERY_LAZY_PLUGINS_TIMEOUT);

    univer.onDispose(() => {
      worker.terminate();
      (window as any).univer = undefined;
      (window as any).univerAPI = undefined;
    });

    (window as any).univer = univer;
    (window as any).univerAPI = FUniver.newAPI(univer);

    // Add chat functionality to global scope
    try {
      console.log('Attempting to get ChatPlugin from injector...');
      const chatPlugin = injector.get(ChatPlugin);
      if (chatPlugin) {
        console.log('ChatPlugin found!', chatPlugin);
        (window as any).univerChat = {
          show: () => chatPlugin.showChat(),
          hide: () => chatPlugin.hideChat(),
          toggle: () => chatPlugin.toggleChat(),
          sendMessage: (msg: string) => chatPlugin.sendMessage(msg),
          clearChat: () => chatPlugin.clearChat(),
          getMessages: () => chatPlugin.getMessages(),
          // MCP specific methods
          setApiKey: (key: string) => chatPlugin.setApiKey(key),
          hasApiKey: () => chatPlugin.hasApiKey(),
          getSessionId: () => chatPlugin.getSessionId(),
          clearApiKey: () => chatPlugin.clearApiKey()
        };
      } else {
        console.error('ChatPlugin is null or undefined');
      }
    } catch (error) {
      console.warn('ChatPlugin not available:', error);
    }

    customRegisterEvent(univer, (window as any).univerAPI!);
  }

  override render() {
    return html`
      <link rel="stylesheet" href="./main.css" />
      <div style="height: 100%; width: 100%;" id="containerId"></div>
    `;
  }
}

// React wrapper
const SheetApp = createComponent({
  tagName: "my-univer-advanced",
  elementClass: MyUniverAdvanced,
  react: React,
});

render(<SheetApp />, document.getElementById('app')!);
