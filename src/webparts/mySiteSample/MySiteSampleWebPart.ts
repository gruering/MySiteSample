import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './MySiteSampleWebPart.module.scss';
import { SPFx, spfi } from '@pnp/sp';
import { BearerToken } from '@pnp/queryable';
import '@pnp/sp/profiles';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IMySiteSampleWebPartProps {}

export default class MySiteSampleWebPart extends BaseClientSideWebPart<IMySiteSampleWebPartProps> {
   public async render(): Promise<void> {
      const sp = spfi().using(SPFx(this.context));
      const profile = await sp.profiles.myProperties();
      const personalUrl = new URL(profile.PersonalUrl);
      console.log('Personal site is ', personalUrl.href);

      // get a bearer token from the personal site url
      const provider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const token = await provider.getToken(`${personalUrl.protocol}//${personalUrl.hostname}`, true);

      // create new context with personal site token
      const mySp = spfi(personalUrl.href).using(SPFx(this.context), BearerToken(token));

      // check if the list was created, or if it already existed
      const sfMyReservations = await mySp.web.lists.ensure('sfMyReservations', 'Reservations List', 100, false, {  });

      if (sfMyReservations.created) {
         console.log('Reservations List created');
      } else {
         console.log('Reservations List already exists');
      }

      const testItem = await mySp.web.lists.getByTitle("sfMyReservations").items.add( { Title: "Test-Item" });    
      console.log(`Item: ${testItem.data.Id} added`); 
      
      this.domElement.innerHTML = `<div class="${styles.mySiteSample}">Item: ${testItem.data.Id} added</div>`;
   }

   protected onInit(): Promise<void> {
      return super.onInit();
   }

   protected get dataVersion(): Version {
      return Version.parse('1.0');
   }
}
