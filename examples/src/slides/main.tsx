/**
 * Copyright 2023-present DreamNum Co., Ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { createComponent } from '@lit/react';
import { LocaleType, LogLevel, Univer, UniverInstanceType } from '@univerjs/core';
import { render } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverDrawingPlugin } from '@univerjs/drawing';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { DEFAULT_SLIDE_DATA } from '@univerjs/mockdata';
import zhCN from '@univerjs/mockdata/locales/zh-CN';
import { UniverSlidesPlugin } from '@univerjs/slides';
import { UniverSlidesUIPlugin } from '@univerjs/slides-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import { html, LitElement } from 'lit';
import { customElement } from 'lit/decorators.js';
import React from 'react';

import '../global.css';

@customElement('my-univer-slides')
class MySlidesComponent extends LitElement {
    override firstUpdated() {
        const container = this.renderRoot.querySelector('#slidesContainer') as HTMLDivElement;

        const univer = new Univer({
            locale: LocaleType.ZH_CN,
            locales: {
                [LocaleType.ZH_CN]: zhCN,
            },
            logLevel: LogLevel.VERBOSE,
        });

        // core plugins
        univer.registerPlugin(UniverRenderEnginePlugin);
        univer.registerPlugin(UniverUIPlugin, { container });
        univer.registerPlugin(UniverDocsPlugin);
        univer.registerPlugin(UniverDocsUIPlugin);

        // base-render
        univer.registerPlugin(UniverFormulaEnginePlugin);
        univer.registerPlugin(UniverDrawingPlugin);
        univer.registerPlugin(UniverSlidesPlugin);
        univer.registerPlugin(UniverSlidesUIPlugin);

        // khởi tạo slides
        univer.createUnit(UniverInstanceType.UNIVER_SLIDE, DEFAULT_SLIDE_DATA);

        (window as any).univer = univer;
    }

    override render() {
        return html`
            <link rel="stylesheet" href="./main.css" />
            <div style="height: 100%;" id="slidesContainer"></div>
        `;
    }
}

const SlidesApp = createComponent({
    tagName: 'my-univer-slides',
    elementClass: MySlidesComponent,
    react: React,
    events: {},
});

render(<SlidesApp />, document.getElementById('app')!);
