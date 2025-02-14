import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
// import * as Handlebars from 'handlebars';
//import * as HTMLDecoder from 'html-decoder';
//import { HFDepts, DateFormat, cdnrepos } from './enumHelper';
//import { fill } from './helper/lodashHelper';
import * as moment from 'moment';
import { Navigation } from 'spfx-navigation';
import { ItemSize, defaultDateFormat } from './Constants';


const filter: any = require('lodash/filter');
const validator: any = require('validator');

export function getQueryStringParams(key: string) {
    var queryString = new UrlQueryParameterCollection(window.location.href);
    if (queryString) {
        return queryString.getValue(key);
    }
    return '';
}

export function removeURLParameter(parameter: string) {
    //prefer to use l.search if you have a location/link object
    var url = window.location.href;
    var urlparts = url.split('?');
    if (urlparts.length >= 2) {
        var prefix = encodeURIComponent(parameter) + '=';
        var pars = urlparts[1].split(/[&;]/g);
        //reverse iteration as may be destructive
        for (var i = pars.length; i-- > 0;) {
            //idiom for string.startsWith
            if (pars[i].lastIndexOf(prefix, 0) !== -1) {
                pars.splice(i, 1);
            }
        }
        return urlparts[0] + (pars.length > 0 ? '?' + pars.join('&') : '');
    }
    return url;
}

export function removeHashValue() {
    var url = window.location.href;
    return url.split('#')[0];
}

export function replaceURL(url: string) {
    window.history.replaceState({}, '', url);
}

// export function getTemplateValue(template: string, value: any) {
//     const hTemplate = Handlebars.compile(HTMLDecoder.decode(template));
//     return HTMLDecoder.decode(hTemplate(value));
// }

export function formatDate(datetime: string, defaultFormat: boolean, dateformat?: string) {
    if (datetime) {
        if (defaultFormat) {
            return moment(datetime).format(defaultDateFormat);
        } else {
            return moment(datetime).format(dateformat);
        }
    }
}

// export function formatDateBasedOnSourceFormat(datetime: string, sourceFormat: string, dateFormat?: string) {
//     if (datetime) {
//         if (dateFormat && dateFormat.length > 0) {
//             return moment(datetime, sourceFormat).format(dateFormat);
//         } else return moment(datetime, sourceFormat).format(DateFormat);
//     }
// }

export function dateDiff(date1: string, date2: string, currentDate: boolean) {
    if (currentDate) {
        return moment(date1).diff(moment(), 'days');
    } else return moment(date1).diff(moment(date2), 'days');
}

// export function getPreviewURL(siteurl, siteid, webid, fileid, resolution) {
//     return `${siteurl}/_layouts/15/getpreview.ashx?resolution=${resolution}&guidSite=${siteid}&guidWeb=${webid}&guidFile=${fileid}&clientType=modernWebPart`;
// }

// export function checkForHFUser(userDept: string) {
//     let redirectDepts = filter(HFDepts, (o) => { return o.toLowerCase() == userDept.toLowerCase(); });
//     if (redirectDepts.length > 0) {
//         return true;
//     }
//     return false;
// }

export function getGuid(): string {
    return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
}

export function s4(): string {
    return Math.floor((1 + Math.random()) * 0x10000)
        .toString(16)
        .substring(1);
}

export function isValidURL(url: string): boolean {
    return url ? validator.isURL(url.trim()) : false;
}

export function getValueFromArray(arr: any[], key: string, valToCheck: string, returnKey: string): any {
    if (arr && arr.length > 0) {
        let fil: any[] = filter(arr, (o: any) => { return o[key].toLowerCase() == valToCheck.toLowerCase(); });
        if (fil && fil.length > 0) {
            return fil[0][returnKey];
        }
    }
    return '';
}

export function hideDefaultPageTitle(): void {
    let style: string;
    style = `
			div[class*="pageTitle_"]
			{
				display: none !important;
			}
		`;
    var head = document.head || document.getElementsByTagName('head')[0];
    var styletag = document.createElement('style');
    styletag.type = 'text/css';
    styletag.appendChild(document.createTextNode(style));
    head.appendChild(styletag);
}

export function getFirstImageContent(divContent: any): string {
    // let retImgUrl: string = '';
    // let firstImage = $(divContent).find('img:first');
    // if (firstImage.length > 0) retImgUrl = firstImage.attr('src');
    let regexImage = /<img[^>]+src="([^">]+)/g;
    var m, urls = [];
    while (m = regexImage.exec(divContent)) {
        urls.push(m[1]);
    }
    return urls[0];
}

export function addCSS(cssStyles: string): void {
    let style: string = cssStyles;
    var head = document.head || document.getElementsByTagName('head')[0];
    var styletag = document.createElement('style');
    styletag.type = 'text/css';
    styletag.appendChild(document.createTextNode(style));
    head.appendChild(styletag);
}

export function hideControls(classnames: string): void {
    let style: string;
    style = `
			${classnames}
			{
				display: none !important;
         }
         .CanvasSection--read .ControlZone {
            margin-top: 0px;
         }
		`;
    var head = document.head || document.getElementsByTagName('head')[0];
    var styletag = document.createElement('style');
    styletag.type = 'text/css';
    styletag.appendChild(document.createTextNode(style));
    head.appendChild(styletag);
}

export function openURL(url: string, newTab: boolean): void {
    // if (newTab) window.open(url, newTab ? '_blank' : '');
    // else window.location.href = url;
    if (newTab) window.open(url, newTab ? '_blank' : '');
    else window.location.href = url;
    //Navigation.navigate(url, false);
}

// export function getRandomArray(arrayLength: number): any[] {
//     return fill([...Array(arrayLength)], 'a', 0, arrayLength);
// }

export function getTileWidth(wpwidth: number, itemSize: ItemSize): string {
    if (wpwidth < 300) {
        switch (itemSize) {
            case ItemSize.Small:
                return '30%';
            case ItemSize.Medium:
                return '43%';
            case ItemSize.Large:
                return '65%';
            case ItemSize['Extra Large']:
                return '80%';
        }
    } else if (wpwidth >= 300 && wpwidth <= 400) {
        switch (itemSize) {
            case ItemSize.Small:
                return '27%';
            case ItemSize.Medium:
                return '43%';
            case ItemSize.Large:
                return '65%';
            case ItemSize['Extra Large']:
                return '80%';
        }
    } else if (wpwidth > 400 && wpwidth <= 600) {
        switch (itemSize) {
            case ItemSize.Small:
                return '25%';
            case ItemSize.Medium:
                return '40%';
            case ItemSize.Large:
                return '60%';
            case ItemSize['Extra Large']:
                return '80%';
        }
    } else if (wpwidth > 600 && wpwidth <= 1100) {
        switch (itemSize) {
            case ItemSize.Small:
                return '16%';
            case ItemSize.Medium:
                return '22%';
            case ItemSize.Large:
                return '28%';
            case ItemSize['Extra Large']:
                return '33%';
        }
    } else if (wpwidth > 1100) {
        switch (itemSize) {
            case ItemSize.Small:
                return '11%';
            case ItemSize.Medium:
                return '15%';
            case ItemSize.Large:
                return '20%';
            case ItemSize['Extra Large']:
                return '33%';
        }
    }
}

export function getTileWidth1(tilesCount: number): string {
    if (tilesCount == 3)
        return `28.7%`;
    else if (tilesCount == 4)
        return `21.7%`;
    else if (tilesCount == 5)
        return `17.7%`;
    else if (tilesCount == 6)
        return `14.1%`;
    else if (tilesCount == 7)
        return `12.9%`;
    else if (tilesCount == 8)
        return `11.1%`;
    else if (tilesCount == 9)
        return `9.7%`;
}

export function _getBoxStyleItemWidth(wpZoneWidthDynamic: number, wpZoneWidth: number): string {
    //console.log("Webpart width: ", props.wpZoneWidth, props.wpZoneWidthDynamic);
    let finalWidth: number = wpZoneWidthDynamic && wpZoneWidthDynamic !== 0 ? wpZoneWidthDynamic : wpZoneWidth;
    if (finalWidth <= 400) return '375px';
    else if (finalWidth >= 550 && finalWidth <= 749) return '276px'; //283px
    else if (finalWidth >= 750 && finalWidth <= 1099) return '250px';
    else if (finalWidth >= 1100) return '290px';
    else return '270px';
}

export function _isSearchPositionAvailable(wpZoneWidthDynamic: number, wpZoneWidth: number): boolean {
    let finalWidth: number = wpZoneWidthDynamic && wpZoneWidthDynamic !== 0 ? wpZoneWidthDynamic : wpZoneWidth;
    if (finalWidth < 749) return false;
    else return true;
}

// export function getFileCDNUrl(publiccdnurl: string, portalurl: string, weburl: string, fileurl: string) {
//     let returl: string = '';
//     let cdnFilter = cdnrepos.filter(o => { return fileurl.toLowerCase().indexOf(o.toLowerCase()) >= 0 && fileurl.toLowerCase().indexOf('/lists/') <= 0; });
//     if (cdnFilter && cdnFilter.length > 0) {
//         if (fileurl.toLowerCase().indexOf(weburl.toLowerCase()) >= 0) {
//             returl = publiccdnurl + weburl.replace("https://", "/") + "/" + fileurl.replace(weburl, '');
//         } else {
//             returl = publiccdnurl + "/" + portalurl.replace("https://", "") + fileurl;
//         }
//     } else returl = fileurl;
//     return returl;
// }

// export function getFileCDNUrl1(publiccdnurl: string, webserurl: string, weburl: string, fileurl: string) {
//     let returl: string = '';
//     if (fileurl && fileurl.length > 0) {
//         let cdnFilter = cdnrepos.filter(o => { return fileurl.toLowerCase().indexOf(o.toLowerCase()) >= 0 && fileurl.toLowerCase().indexOf('/lists/') <= 0; });
//         if (cdnFilter && cdnFilter.length > 0) {
//             if (fileurl.toLowerCase().indexOf(weburl.toLowerCase()) >= 0) {
//                 returl = publiccdnurl + weburl.replace("https://", "/") + "/" + fileurl.replace(weburl, '');
//             } else {
//                 let portalurl = webserurl == "/" ? weburl.replace("https://", "") : weburl.replace(webserurl, "").replace("https://", "");
//                 returl = publiccdnurl + "/" + portalurl + fileurl;
//             }
//         } else returl = fileurl;
//     }
//     return returl;
// }

// export function UseFullWidth() {
//     const jQuery: any = require('jquery');
//     jQuery("#workbenchPageContent").prop("style", "max-width: none");
//     jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
//     jQuery(".CanvasZone").prop("style", "max-width: none");
// }

export function getFileIcon(filename: string): string {
    let retFileIcon: string = 'Document';
    if (filename && filename.length > 0) {
        let fileExtn: string = filename.split('.').pop();
        switch (fileExtn.toLowerCase()) {
            case 'docx':
            case 'doc':
                retFileIcon = 'WordDocument';
                break;
            case 'xlsx':
            case 'xls':
                retFileIcon = 'ExcelDocument';
                break;
            case 'pptx':
            case 'ppt':
                retFileIcon = 'PowerPointDocument';
                break;
            case 'jpg':
            case 'jpeg':
            case 'png':
            case 'gif':
            case 'tiff':
                retFileIcon = 'FileImage';
                break;
        }
    }
    return retFileIcon;
}