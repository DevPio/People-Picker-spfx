import * as React from "react";
import styles from "./Peoplepicker.module.scss";
import { IPeoplepickerProps } from "./IPeoplepickerProps";
import { NormalPeoplePicker, IPersonaProps, Label } from "office-ui-fabric-react";

const Peoplepicker: React.FC<IPeoplepickerProps> = (props) => {
  const [people, setPeople] = React.useState<any>([]);

  const searchPeople = (
    terms: string
  ): IPersonaProps[] | Promise<IPersonaProps[]> =>
    new Promise((resolve, reject) => {
      fetch(
        `https://solucoesspfx.sharepoint.com/_api/search/query?querytext='*${terms}*'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        }
      )
        .then((response) => {
          return response.json();
        })
        .then((response: { PrimaryQueryResult: any }) => {
          let relevantResults: any =
            response.PrimaryQueryResult.RelevantResults;
          let resultCount: number = relevantResults.TotalRows;
          let people = [];

          if (resultCount > 0) {
            relevantResults.Table.Rows.forEach(function (row) {
              let persona: IPersonaProps = {};
              row.Cells.forEach(function (cell) {
                if (cell.Key === "PictureURL") persona.imageUrl = cell.Value;
                if (cell.Key === "PreferredName")
                  persona.primaryText = cell.Value;
                if (cell.Key === "WorkEmail")
                  persona.secondaryText = cell.Value;
              });
              people.push(persona);
            });
          }
          resolve(people);
        })
        .catch((error) => {
          reject(error);
        });
    });

  const getPeople = async (filterText: string) => {
    if (filterText) {
      if (filterText.length > 2) {
        const search = await searchPeople(filterText);
        setPeople(search);
        return searchPeople(filterText);
      }
    } else {
      return [];
    }
  };

  return (
    <div className={styles.peoplepicker}>
      <Label>Peopler Picker</Label>
      <NormalPeoplePicker onResolveSuggestions={getPeople} />
    </div>
  );
};

export default Peoplepicker;
