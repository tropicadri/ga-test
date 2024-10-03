import { Compat, feature } from "compute-baseline/browser-compat-data";
import { getStatus } from "compute-baseline";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import webSpecs from "web-specs" assert { type: "json" };
import { Document } from "yaml";
import { google } from "googleapis";

import { features } from "../index.js";

const clientEmail = 'baseline@baseline-test-431120.iam.gserviceaccount.com';
const privateKey = "-----BEGIN PRIVATE KEY-----\nMIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQC+wsgdk/KuCkhh\nW+5JVRdrlrNYCR44TselfFNDQKr9itPBALiXgOrcltHryvzNm0fGxmKQV+T/RE9e\nblkW4GeqQZ+jZB6YPNe8XgcpHNLUlHenSj/fk8crJJDG13GAe7VXYZxsK4VETEzO\nH1yIdx7xfbdAA2QdzlENOEnQ6IcvMR7u2A2en90H9ADB5zwJwLTNXLSWqDJrRy5T\n1f5NxnvWSva/vODgKYqta3ffm72gluQJL523LjKNxujImumzAoP4TvRqBAFj0WGM\nopK1x7kLH9Qbi99UKwY0TUjE6aFdmDS4K+IulgOi273fAhEqocnvJSDjd/lb4hkT\nIkrYQWnvAgMBAAECggEATSbmXOAbhLxv7cCKrWoW4NLAkliNumRSZuHTox2iioz9\ne8/Y2J2xX6o4GM+uU1H8Q6606oBCA+QS3bRvn8r+4mNMN38kyJnC/1JgPwaDVqr9\ngdGL59cS9KYDUAiOxVJDU6YM58rqSWc6oRsaB691+mHIHcHBWjS5n0kmwbFqC3F1\nXCZTsqsKjhDf9NAdAxIFFaDgGVr/beDlCq1E+bHIIiSt0MD+a7gkyeLGNK39bgXG\n9NN2x7gGkhxSX3d+E31/P6mkc9NmnqxvETjFU5VOnVaF1M6GCls113BowbzS0dPe\nG68k2a+PcX/M1VvHv311Fgdtna1C6FO9RO20iZK9QQKBgQDrmKZPXhW7n/6QFHqr\nex2qR7TPSfgTsyaKtpQSIaerRLmuMMkVxuIemCAiDF0PrycCxGLPQjukdWqoSBK+\n0idKnoysF1iTG49sDDFxQS4NAC0nTzHADGsZElYeynYozjNxJp52eIcMf+MVtXDO\nd7Fh90oGSAmHia4aIke4GNc58QKBgQDPSBisBvNG7VgH3qkuYUCzLUO0Sqsg+IZy\nZnkmxNkByZizibqk50NwKJ1o4FvYnOWalZO6+Xgp5ld08rZbfPm3qyZIRSZMSOKJ\n38oE9dwAqV6zQaM+7msu9OXfuYhinEVgHOOCOd6WFLQtVm72mXfsnmWRMvXADnHU\n3QGnykwB3wKBgQCwEaro0bWIPO/wGX37vczTvzcfW3ZssoxxSjWqv0AhPQFPbuI4\nxiOjHamWCItb91g+NyppHmSYaa/GB+cKeOAoYglbay8k8YnkA9DHhSC3UpRsNkFF\n94HqgMufw0s+/6tuOqBehYDYFJNPXAPsKYOFZVwLCbcldf2JRmk26kme4QKBgQCu\nQJetFnKxb5mPlBl5ifxQJH9EM3lFUGgGCCtU8irUG9upIuGq+1dwFFSv89yCbvN+\nTrxQdKQV5TQkdl3i718mjlSycJ9Phf4blexI2z7Ft5yK6WUI5HlqkiolN911uJWB\nxQdlVKzQUjtd2cmSuVVdr5MGy2ZYvEvpfJmFqsTW1QKBgQCrMQld5N7rsj4AZh7N\n0g+HJZXQ5rNKRA71XT2flY7RMFw3CYPwMp9hzUi6Sdyhj6F9YXX/ExjPL5G7AkjG\nX3fYpTkDaDUTA95yaHkq6Z+G2ATf1J534y7zpxP/2yy6z9fMIm4E0fqNgSHg+9l3\nCUqhpjex0vSiiCKhBOC+2x9C2g==\n-----END PRIVATE KEY-----\n"
const googleSheetId = '1BjPgmfHQl0dyd8Ung9ql6vs8PkJWztAxevMkchvAGSo';
const googleSheetPage = 'Sheet1';

// authenticate the service account
const googleAuth = new google.auth.JWT(
  clientEmail,
  null,
  privateKey.replace(/\\n/g, '\n'),
  'https://www.googleapis.com/auth/spreadsheets'
);

async function updateSheet(data) {
  try {
    // google sheet instance
    const sheetInstance = await google.sheets({ version: 'v4', auth: googleAuth});

    const updateToGsheet = [
      [ 'Seoul', '9,776,000', 'Asia', 'KRW' ],
      [ 'Copenhagen', '602,000', 'Europe', 'EUR' ],
      [ 'Amsterdam', '821,000', 'Europe', 'EUR' ],
      [ 'Helsinki', '631,000', 'Europe', 'EUR' ],
      [ 'Sydney', '5,312,000', 'Oceania', 'AUD' ]
    ];
    
    // update data in the range
    await sheetInstance.spreadsheets.values.update({
        auth: googleAuth,
        spreadsheetId: googleSheetId,
        range: `${googleSheetPage}!A2:C${data.length + 1}`,
        valueInputOption: 'RAW',
        resource: {
          values: data,
        },
    });
  }
  catch(err) {
    console.log("updateSheet func() error", err);  
  }
}



type WebSpecsSpec = (typeof webSpecs)[number];

function* getPages(spec): Generator<string> {
  yield spec.url;
  if (spec.nightly?.url) {
    yield spec.nightly.url;
  }
  if (spec.nightly?.pages) {
    yield* spec.nightly.pages;
  }
}

function normalize(page: string) {
  const url = new URL(page);
  // Remove any fragment.
  url.hash = "";
  // Collapse HTML and ECMA-262 multipage into a single canonical page.
  const multipageIndex = url.pathname.indexOf("/multipage/");
  if (multipageIndex !== -1) {
    url.pathname = url.pathname.substring(0, multipageIndex + 1);
  }
  // Strip levels from CSS specs.
  if (url.hostname.startsWith("drafts.")) {
    url.pathname = url.pathname.replace(/-\d+\/$/, "/");
  }
  return String(url);
}

function formatIdentifier(s: string): string {
  return s
    .toLowerCase()
    .split(/[^a-z0-9-]+/)
    .join("-");
}

async function main() {
  const compat = new Compat();

  // Build a map of used BCD keys to feature.
  const webFeatures = new Map<string, string>();
  Object.values(features).map((data) => {
    if (data.compat_features) {
      for (const compatFeature of data.compat_features) {
        webFeatures.set(compatFeature, data.name);
      }
    }
  });

  // Build a map from URLs to spec.
  const pageToSpec = new Map<string, WebSpecsSpec>();
  for (const spec of webSpecs) {
    for (const page of getPages(spec)) {
      pageToSpec.set(normalize(page), spec);
    }
  }

  // Iterate BCD and group compat features by spec.
  const specToCompatFeatures = new Map<WebSpecsSpec, Set<string>>();
  let count = 0;
  const data = [];
  for (const feature of compat.walk()) {
    console.log('feature', feature.id);
    // Skip deprecated and non-standard features.
    if (feature.deprecated || !feature.standard_track) {
      continue;
    }

    try{
      const status = getStatus('sample', feature.id, compat);
      const isB = status.baseline_low_date ? true : false;

      const row = [feature.id, isB, status.baseline_low_date];
  
      data.push(row);
    }catch(e){
      console.log('error', e);
      continue;
    }

  
    
    
   



    

  }

  updateSheet(data);

  // // Separate out features that are already part of web-features.
  // for (const [spec, compatFeatures] of specToCompatFeatures.entries()) {
  //   const usedFeatures = new Map<string, Set<String>>();
  //   for (const key of compatFeatures) {
  //     if (webFeatures.has(key)) {
  //       const feature = webFeatures.get(key);
  //       if (usedFeatures.has(feature)) {
  //         usedFeatures.get(feature).add(key);
  //       } else {
  //         usedFeatures.set(feature, new Set([key]));
  //       }
  //       compatFeatures.delete(key);
  //     }
  //   }

  //   // If all features are already part of web-features, skip this spec.
  //   if (compatFeatures.size === 0) {
  //     continue;
  //   }

  //   // Write out draft feature per spec.
  //   const id = formatIdentifier(spec.shortname);

  //   const feature = new Document({
  //     draft_date: new Date().toISOString().substring(0, 10),
  //     name: spec.title,
  //     description: "TODO",
  //     spec: spec.nightly?.url ?? spec.url,
  //     compat_features: Array.from(compatFeatures).sort(),
  //   });

  //   if (usedFeatures.size > 0) {
  //     let usedFeaturesComment = ` The following features in the spec are already part of web-features:\n`;
  //     for (const [feature, keys] of usedFeatures.entries()) {
  //       usedFeaturesComment += ` - ${feature}:\n   - ${Array.from(keys).join("\n   - ")}\n`;
  //     }

  //     feature.comment = usedFeaturesComment.trimEnd();
  //   }
  //   await fs.writeFile(`features/draft/spec/${id}.yml`, feature.toString());
  // }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  await main();
}
