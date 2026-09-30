import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import * as strings from 'PlanMyDayApplicationCustomizerStrings';
import { App } from './components/App';

const LOG_SOURCE: string = 'PlanMyDayApplicationCustomizer';

// No custom properties required; the extension renders itself in the Top placeholder.
export type IPlanMyDayApplicationCustomizerProperties = Record<string, never>;

/** Renders the "Plan My Day" launcher bar and slide-out panel in the page's Top placeholder. */
export default class PlanMyDayApplicationCustomizer
  extends BaseApplicationCustomizer<IPlanMyDayApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders);
    this._renderPlaceholders();

    return Promise.resolve();
  }

  private _renderPlaceholders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    }

    if (!this._topPlaceholder || !this._topPlaceholder.domElement) {
      Log.error(LOG_SOURCE, new Error('The Top placeholder is not available on this page.'));
      return;
    }

    const element: React.ReactElement = React.createElement(App, { context: this.context });
    ReactDOM.render(element, this._topPlaceholder.domElement);
  }

  protected onDispose(): void {
    ReactDOM.unmountComponentAtNode(this._topPlaceholder?.domElement as Element);
    super.onDispose();
  }
}

