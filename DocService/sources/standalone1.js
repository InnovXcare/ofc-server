var os = require('os');
var path = require('path');
var fs = require('fs');
var co = require('co');
var config = require('config');
const bytes = require('bytes');

const express = require('express');
var spawnAsync = require('@expo/spawn-async');
var commonDefines = require('./commondefines');
const utilsDocService = require("./utilsDocService");
var taskResult = require('./taskresult');
const formatChecker = require('./../../Common/sources/formatchecker');
const operationContext = require('./../../Common/sources/operationContext');
var utils = require('./../../Common/sources/utils');
// var storage = require('./../../Common/sources/storage/storage-base');
// var utils = require('./../../Common/sources/utils');
var constants = require('./constants');

var TEMP_PREFIX = 'FILE_CONVERT';
const cfgFontDir = config.get('FileConverter.converter.fontDir');
const cfgX2tPath = config.get('FileConverter.converter.x2tPath');
const cfgPresentationThemesDir = config.get('FileConverter.converter.presentationThemesDir');
const cfgDocbuilderPath = config.get('FileConverter.converter.docbuilderPath');
const cfgArgs = config.get('FileConverter.converter.args');
const cfgInputLimits = config.get('FileConverter.converter.inputLimits');
const cfgRequesFilteringAgent = config.util.cloneDeep(config.get('services.CoAuthoring.request-filtering-agent'));
const cfgExternalRequestDirectIfIn = config.get('externalRequest.directIfIn');
const cfgExternalRequestAction = config.get('externalRequest.action');
const cfgSpawnOptions = config.util.cloneDeep(config.get('FileConverter.converter.spawnOptions'));

let inputLimitsXmlCache;

const router = express.Router();

router.get('/', (req, res) => {
    res.send('standalone1 ok');
});

router.get('/convert', (req, res) => {
    return co(function* () {
    let ctx = new operationContext.Context();
    try {
      ctx.initFromRequest(req);
      yield ctx.initTenantCache();
      ctx.logger.info('convertRequest start');
      let params = {
        async: true,
        url: path.join(__dirname, '../../samples/test.docx'),
        outputtype: 'bin',
        filetype: 'docx',
        title: 'test file',
        key: 'test111',
        password: null,
        region: 'en',
        // pdf: '',
        // codePage: '',
        // delimiter: '',
        // delimiterChar: '',
        // documentLayout: '',
        // spreadsheetLayout: '',
        // watermark: '',
        // thumbnail: '',
        // documentRenderer: '',
      };

      // const params = { // write all the conversion parameters to the params dictionary
      //   async,
      //   url: documentUri,
      //   outputtype: toExtension.replace('.', ''),
      //   filetype: fromExt.replace('.', ''),
      //   title,
      //   key: revisionId,
      //   password: filePass,
      //   region: lang,
      // };

      // let authRes = yield docsCoServer.getRequestParams(ctx, req);
      // if(authRes.code === constants.NO_ERROR){
      //   params = authRes.params;
      // } else {
      //   ctx.logger.warn('convertRequest auth failed %j', authRes);
      //   utils.fillResponse(req, res, new commonDefines.ConvertStatus(authRes.code), isJson);
      //   return;
      // }
      let filetype = params.filetype || params.fileType || '';
      let outputtype = params.outputtype || params.outputType || '';
      ctx.setDocId(params.key);
      const isJson = true;

      // if (params.key && !constants.DOC_ID_REGEX.test(params.key)) {
      //   ctx.logger.warn('convertRequest unexpected key = %s', params.key);
      //   utils.fillResponse(req, res, new commonDefines.ConvertStatus(constants.CONVERT_PARAMS), isJson);
      //   return;
      // }
      if (filetype && !constants.EXTENTION_REGEX.test(filetype)) {
        ctx.logger.warn('convertRequest unexpected filetype = %s', filetype);
        utils.fillResponse(req, res, new commonDefines.ConvertStatus(constants.CONVERT_PARAMS), isJson);
        return;
      }
      let outputFormat = outputtype === 'bin' ? constants.AVS_OFFICESTUDIO_FILE_CANVAS_WORD :formatChecker.getFormatFromString(outputtype);
      if (constants.AVS_OFFICESTUDIO_FILE_UNKNOWN === outputFormat) {
        ctx.logger.warn('convertRequest unexpected outputtype = %s', outputtype);
        utils.fillResponse(req, res, new commonDefines.ConvertStatus(constants.CONVERT_PARAMS), isJson);
        return;
      }
      // let oformAsPdf;
      // if (params.pdf) {
      //   if (true === params.pdf.pdfa && constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDF === outputFormat) {
      //     outputFormat = constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDFA;
      //   } else if (false === params.pdf.pdfa && constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDFA === outputFormat) {
      //     outputFormat = constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDF;
      //   }
      //   if (params.pdf.form && (constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDF === outputFormat ||
      //     constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDFA === outputFormat)) {
      //     outputFormat = constants.AVS_OFFICESTUDIO_FILE_DOCUMENT_OFORM_PDF;
      //   } else if (false === params.pdf.form) {
      //     oformAsPdf = true;
      //   }
      // }
      //todo use hash of params as id
      let docId = 'conv_' + params.key + '_' + outputFormat;
      var cmd = new commonDefines.InputCommand();
      cmd.setCommand('conv');
      cmd.setUrl(params.url);
      cmd.setEmbeddedFonts(false);//params.embeddedfonts'];
      cmd.setFormat(filetype);
      cmd.setDocId(docId);
      cmd.setOutputFormat(outputFormat);
      // cmd.setOformAsPdf(oformAsPdf);
      let outputExt = formatChecker.getStringFromFormat(cmd.getOutputFormat());

      // cmd.setCodepage(commonDefines.c_oAscEncodingsMap[params.codePage] || commonDefines.c_oAscCodePageUtf8);
      // cmd.setDelimiter(parseIntParam(params.delimiter) || commonDefines.c_oAscCsvDelimiter.Comma);
      // if(undefined != params.delimiterChar)
      //   cmd.setDelimiterChar(params.delimiterChar);
      if (params.region) {
        cmd.setLCID(utilsDocService.localeToLCID(params.region));
      }
      // let jsonParams = {};
      // if (params.documentLayout) {
      //   jsonParams['documentLayout'] = params.documentLayout;
      // }
      // if (params.spreadsheetLayout) {
      //   jsonParams['spreadsheetLayout'] = params.spreadsheetLayout;
      // }
      // if (params.watermark) {
      //   jsonParams['watermark'] = params.watermark;
      // }
      // if (Object.keys(jsonParams).length > 0) {
      //   cmd.appendJsonParams(jsonParams);
      // }
      // if (params.password) {
      //   if (params.password.length > constants.PASSWORD_MAX_LENGTH) {
      //     ctx.logger.warn('convertRequest password too long actual = %s; max = %s', params.password.length, constants.PASSWORD_MAX_LENGTH);
      //     utils.fillResponse(req, res, new commonDefines.ConvertStatus(constants.CONVERT_PARAMS), isJson);
      //     return;
      //   }
      //   let encryptedPassword = yield utils.encryptPassword(ctx, params.password);
      //   cmd.setPassword(encryptedPassword);
      // }
      // if (authRes.isDecoded) {
      //   cmd.setWithAuthorization(true);
      // }
      cmd.setWithAuthorization(true);

      // var thumbnail = params.thumbnail;
      // if (thumbnail) {
      //   if (typeof thumbnail === 'string') {
      //     thumbnail = JSON.parse(thumbnail);
      //   }
      //   var thumbnailData = new commonDefines.CThumbnailData(thumbnail);
      //   //constants from CXIMAGE_FORMAT_
      //   switch (cmd.getOutputFormat()) {
      //     case constants.AVS_OFFICESTUDIO_FILE_IMAGE_JPG:
      //       thumbnailData.setFormat(3);
      //       break;
      //     case constants.AVS_OFFICESTUDIO_FILE_IMAGE_PNG:
      //       thumbnailData.setFormat(4);
      //       break;
      //     case constants.AVS_OFFICESTUDIO_FILE_IMAGE_GIF:
      //       thumbnailData.setFormat(2);
      //       break;
      //     case constants.AVS_OFFICESTUDIO_FILE_IMAGE_BMP:
      //       thumbnailData.setFormat(1);
      //       break;
      //   }
      //   cmd.setThumbnail(thumbnailData);
      //   if (false === thumbnailData.getFirst() && 0 !== (constants.AVS_OFFICESTUDIO_FILE_IMAGE & cmd.getOutputFormat())) {
      //     outputExt = 'zip';
      //   }
      // }
      // var documentRenderer = params.documentRenderer;
      // if (documentRenderer) {
      //   if (typeof documentRenderer === 'string') {
      //     documentRenderer = JSON.parse(documentRenderer);
      //   }
      //   var textParamsData = new commonDefines.CTextParams();
      //   switch (documentRenderer.textAssociation) {
      //     case 'plainParagraph':
      //       textParamsData.setAssociation(3);
      //       break;
      //     case 'plainLine':
      //       textParamsData.setAssociation(2);
      //       break;
      //     case 'blockLine':
      //       textParamsData.setAssociation(1);
      //       break;
      //     case 'blockChar':
      //     default:
      //       textParamsData.setAssociation(0);
      //       break;
      //   }
      //   cmd.setTextParams(textParamsData);
      // }
      if (params.title) {
        cmd.setTitle(path.basename(params.title, path.extname(params.title)) + '.' + outputExt);
      }
      var async = (typeof params.async === 'string') ? 'true' == params.async : params.async;
      // if (async && !req.query[constants.SHARD_KEY_API_NAME] && !req.query[constants.SHARD_KEY_WOPI_NAME] && process.env.DEFAULT_SHARD_KEY) {
      //   ctx.logger.warn('convertRequest set async=false. Pass query string parameter "%s" to correctly process request in sharded cluster', constants.SHARD_KEY_API_NAME);
      //   async = false;
      // }
      if (constants.AVS_OFFICESTUDIO_FILE_UNKNOWN !== cmd.getOutputFormat()) {
        let fileTo = constants.OUTPUT_NAME + '.' + outputExt;
        // var status = yield* convertByCmd(ctx, cmd, async, fileTo, undefined, undefined, undefined, undefined, true);
        const resData = yield* ExecuteTask(ctx, {
          cmd,
          fileTo
        });
        // if (status.end) {
        //   let fileToPath = yield* getConvertPath(ctx, docId, fileTo, cmd.getOutputFormat());
        //   status.setExtName(path.extname(fileToPath));
        //   status.setUrl(yield* getConvertUrl(ctx, utils.getBaseUrlByRequest(ctx, req), fileToPath, cmd.getTitle()));
        //   ctx.logger.debug('convertRequest: url = %s', status.url);
        // }
        console.log('from execute task::', resData);
        var status = new commonDefines.ConvertStatus(constants.NO_ERROR);
        status.end = true;
        const isJson = true;
        utils.fillResponse(req, res, status, isJson);
      } else {
        var addresses = utils.forwarded(req);
        ctx.logger.warn('Error convert unknown outputtype: query = %j from = %s', params, addresses);
        utils.fillResponse(req, res, new commonDefines.ConvertStatus(constants.UNKNOWN), isJson);
      }
    } catch (e) {
      ctx.logger.error('convertRequest error: %s', e.stack);
      utils.fillResponse(req, res, new commonDefines.ConvertStatus(constants.UNKNOWN), isJson);
    } finally {
      ctx.logger.info('convertRequest end');
    }
  });
});


function TaskQueueDataConvert(ctx, execObj) {
  // var cmd = task.getCmd();
  let {cmd} = execObj;
  this.key = cmd.getDocId();
  if (cmd.getSaveKey()) {
    this.key += cmd.getSaveKey();
  }
  this.fileFrom = null;
  this.fileTo = null;
  this.title = cmd.getTitle();
  if(constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDFA !== cmd.getOutputFormat()){
    this.formatTo = cmd.getOutputFormat();
  } else {
    this.formatTo = constants.AVS_OFFICESTUDIO_FILE_CROSSPLATFORM_PDF;
    this.isPDFA = true;
  }
  this.csvTxtEncoding = cmd.getCodepage();
  this.csvDelimiter = cmd.getDelimiter();
  this.csvDelimiterChar = cmd.getDelimiterChar();
  // this.paid = task.getPaid();
  this.paid = true;
  this.embeddedFonts = cmd.embeddedfonts;
  // this.fromChanges = task.getFromChanges();
  //todo
  const tenFontDir = ctx.getCfg('FileConverter.converter.fontDir', cfgFontDir);
  if (tenFontDir) {
    this.fontDir = path.resolve(tenFontDir);
  } else {
    this.fontDir = null;
  }
  const tenPresentationThemesDir = ctx.getCfg('FileConverter.converter.presentationThemesDir', cfgPresentationThemesDir);
  this.themeDir = path.resolve(tenPresentationThemesDir);
  this.mailMergeSend = cmd.mailmergesend;
  this.thumbnail = cmd.thumbnail;
  this.textParams = cmd.getTextParams();
  this.jsonParams = JSON.stringify(cmd.getJsonParams());
  this.lcid = cmd.getLCID();
  this.password = cmd.getPassword();
  this.savePassword = cmd.getSavePassword();
  this.noBase64 = cmd.getNoBase64();
  this.convertToOrigin = cmd.getConvertToOrigin();
  this.oformAsPdf = cmd.getOformAsPdf();
  this.timestamp = new Date();
}
TaskQueueDataConvert.prototype = {
  serialize: function(ctx, fsPath) {
    let xml = '\ufeff<?xml version="1.0" encoding="utf-8"?>';
    xml += '<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"';
    xml += ' xmlns:xsd="http://www.w3.org/2001/XMLSchema">';
    xml += this.serializeXmlProp('m_sKey', this.key);
    xml += this.serializeXmlProp('m_sFileFrom', this.fileFrom);
    xml += this.serializeXmlProp('m_sFileTo', this.fileTo);
    xml += this.serializeXmlProp('m_sTitle', this.title);
    xml += this.serializeXmlProp('m_nFormatTo', this.formatTo);
    xml += this.serializeXmlProp('m_bIsPDFA', this.isPDFA);
    xml += this.serializeXmlProp('m_nCsvTxtEncoding', this.csvTxtEncoding);
    xml += this.serializeXmlProp('m_nCsvDelimiter', this.csvDelimiter);
    xml += this.serializeXmlProp('m_nCsvDelimiterChar', this.csvDelimiterChar);
    xml += this.serializeXmlProp('m_bPaid', this.paid);
    xml += this.serializeXmlProp('m_bEmbeddedFonts', this.embeddedFonts);
    xml += this.serializeXmlProp('m_bFromChanges', this.fromChanges);
    xml += this.serializeXmlProp('m_sFontDir', this.fontDir);
    xml += this.serializeXmlProp('m_sThemeDir', this.themeDir);
    if (this.mailMergeSend) {
      xml += this.serializeMailMerge(this.mailMergeSend);
    }
    if (this.thumbnail) {
      xml += this.serializeThumbnail(this.thumbnail);
    }
    if (this.textParams) {
      xml += this.serializeTextParams(this.textParams);
    }
    xml += this.serializeXmlProp('m_sJsonParams', this.jsonParams);
    xml += this.serializeXmlProp('m_nLcid', this.lcid);
    xml += this.serializeXmlProp('m_oTimestamp', this.timestamp.toISOString());
    xml += this.serializeXmlProp('m_bIsNoBase64', this.noBase64);
    xml += this.serializeXmlProp('m_sConvertToOrigin', this.convertToOrigin);
    xml += this.serializeLimit(ctx);
    xml += this.serializeOptions(ctx, false, this.oformAsPdf);
    xml += '</TaskQueueDataConvert>';
    fs.writeFileSync(fsPath, xml, {encoding: 'utf8'});
  },
  serializeHidden: function(ctx) {
    var t = this;
    return co(function* () {
      let xml;
      if (t.password || t.savePassword) {
        xml = '<TaskQueueDataConvert>';
        if(t.password) {
          let password = yield utils.decryptPassword(ctx, t.password);
          xml += t.serializeXmlProp('m_sPassword', password);
        }
        if(t.savePassword) {
          let savePassword = yield utils.decryptPassword(ctx, t.savePassword);
          xml += t.serializeXmlProp('m_sSavePassword', savePassword);
        }
        xml += '</TaskQueueDataConvert>';
      }
      return xml;
    });
  },
  serializeOptions: function (ctx, isInJwtToken, oformAsPdf) {
    const tenRequesFilteringAgent = ctx.getCfg('services.CoAuthoring.request-filtering-agent', cfgRequesFilteringAgent);
    const tenExternalRequestDirectIfIn = ctx.getCfg('externalRequest.directIfIn', cfgExternalRequestDirectIfIn);
    const tenExternalRequestAction = ctx.getCfg('externalRequest.action', cfgExternalRequestAction);
    let allowList = tenExternalRequestDirectIfIn.allowList;
    let allowNetworkRequest = tenExternalRequestAction.allow;
    let allowPrivateIP = !tenExternalRequestAction.blockPrivateIP && tenRequesFilteringAgent.allowPrivateIPAddress;
    let proxyUrl = tenExternalRequestAction.proxyUrl;
    let proxyUser = tenExternalRequestAction.proxyUser;
    let proxyHeaders = tenExternalRequestAction.proxyHeaders;
    if (allowList.length === 0 && tenExternalRequestDirectIfIn.jwtToken && isInJwtToken) {
      allowNetworkRequest = true;
      allowPrivateIP = true;
      proxyUrl = "";
      proxyUser = null;
      proxyHeaders = {};
    }
    let xml = "";
    xml += '<options>';
    if (allowList.length > 0) {
      xml += this.serializeXmlProp('allowList', allowList.join(';'));
    }
    xml += this.serializeXmlProp('allowNetworkRequest', allowNetworkRequest);
    xml += this.serializeXmlProp('allowPrivateIP', allowPrivateIP);
    if (proxyUrl) {
      xml += this.serializeXmlProp('proxy', proxyUrl);
    }
    if (proxyUser) {
      let user = proxyUser.username;
      let pass = proxyUser.password;
      xml += this.serializeXmlProp('proxyUser', `${user}:${pass}`);
    }
    let proxyHeadersStr= [];
    for (let name in proxyHeaders) {
      proxyHeadersStr.push(`${name}:${proxyHeaders[name]}`);
    }
    if (proxyHeadersStr.length > 0) {
      xml += this.serializeXmlProp('proxyHeader', proxyHeadersStr.join(';'));
    }
    if (undefined !== oformAsPdf) {
      xml += this.serializeXmlProp('oformAsPdf', oformAsPdf);
    }
    xml += '</options>';
    return xml;
  },
  serializeMailMerge: function(data) {
    var xml = '<m_oMailMergeSend>';
    xml += this.serializeXmlProp('from', data.getFrom());
    xml += this.serializeXmlProp('to', data.getTo());
    xml += this.serializeXmlProp('subject', data.getSubject());
    xml += this.serializeXmlProp('mailFormat', data.getMailFormat());
    xml += this.serializeXmlProp('fileName', data.getFileName());
    xml += this.serializeXmlProp('message', data.getMessage());
    xml += this.serializeXmlProp('recordFrom', data.getRecordFrom());
    xml += this.serializeXmlProp('recordTo', data.getRecordTo());
    xml += this.serializeXmlProp('recordCount', data.getRecordCount());
    xml += this.serializeXmlProp('userid', data.getUserId());
    xml += this.serializeXmlProp('url', data.getUrl());
    xml += '</m_oMailMergeSend>';
    return xml;
  },
  serializeThumbnail: function(data) {
    var xml = '<m_oThumbnail>';
    xml += this.serializeXmlProp('format', data.getFormat());
    xml += this.serializeXmlProp('aspect', data.getAspect());
    xml += this.serializeXmlProp('first', data.getFirst());
    xml += this.serializeXmlProp('width', data.getWidth());
    xml += this.serializeXmlProp('height', data.getHeight());
    xml += '</m_oThumbnail>';
    return xml;
  },
  serializeTextParams: function(data) {
    var xml = '<m_oTextParams>';
    xml += this.serializeXmlProp('m_nTextAssociationType', data.getAssociation());
    xml += '</m_oTextParams>';
    return xml;
  },
  serializeLimit: function(ctx) {
    if (!inputLimitsXmlCache) {
      var xml = '<m_oInputLimits>';
      const tenInputLimits = ctx.getCfg('FileConverter.converter.inputLimits', cfgInputLimits);
      for (let i = 0; i < tenInputLimits.length; ++i) {
        let limit = tenInputLimits[i];
        if (limit.type && limit.zip) {
          xml += '<m_oInputLimit';
          xml += this.serializeXmlAttr('type', limit.type);
          xml += '>';
          xml += '<m_oZip';
          if (limit.zip.compressed) {
            xml += this.serializeXmlAttr('compressed', bytes.parse(limit.zip.compressed));
          }
          if (limit.zip.uncompressed) {
            xml += this.serializeXmlAttr('uncompressed', bytes.parse(limit.zip.uncompressed));
          }
          xml += this.serializeXmlAttr('template', limit.zip.template);
          xml += '/>';
          xml += '</m_oInputLimit>';
        }
      }
      xml += '</m_oInputLimits>';
      inputLimitsXmlCache = xml;
    }
    return inputLimitsXmlCache;
  },
  serializeXmlProp: function(name, value) {
    var xml = '';
    //todo check empty and undefined (password?)
    if (null != value) {
      xml += '<' + name + '>';
      xml += utils.encodeXml(value.toString());
      xml += '</' + name + '>';
    } else {
      xml += '<' + name + ' xsi:nil="true" />';
    }
    return xml;
  },
  serializeXmlAttr: function(name, value) {
    var xml = '';
    if (null != value) {
      xml += ' ' + name + '=\"';
      xml += utils.encodeXml(value.toString());
      xml += '\"';
    }
    return xml;
  }
};

function getTempDir() {
  var tempDir = os.tmpdir();
  var now = new Date();
  var newTemp;
  while (!newTemp || fs.existsSync(newTemp)) {
    var newName = [TEMP_PREFIX, now.getFullYear(), now.getMonth(), now.getDate(),
      '-', (Math.random() * 0x100000000 + 1).toString(36)
    ].join('');
    newTemp = path.join(tempDir, newName);
  }
  fs.mkdirSync(newTemp);
  var sourceDir = path.join(newTemp, 'source');
  fs.mkdirSync(sourceDir);
  var resultDir = path.join(newTemp, 'result');
  fs.mkdirSync(resultDir);
  return {temp: newTemp, source: sourceDir, result: resultDir};
}

function checkPathTraversal(ctx, docId, rootDirectory, filename) {
  if (filename.indexOf('\0') !== -1) {
    console.warn('checkPathTraversal Poison Null Bytes filename=%s', filename);
    return false;
  }
  if (!filename.startsWith(rootDirectory)) {
    console.warn('checkPathTraversal Path Traversal filename=%s', filename);
    return false;
  }
  return true;
};

function downloadFile(ctx, url, dataConvert) {
  const {fileFrom} = dataConvert;
  let res = constants.NO_ERROR;
  // add S3 download logic later
  console.log('fileFrom:::', fileFrom)
  if (!fs.existsSync(fileFrom)) {
    fs.copyFileSync(url, fileFrom);
    // res = constants.CONVERT_DOWNLOAD;
    // console.debug('file not found');
  }
  return res;
}

function uploadFile(ctx, url, dataConvert) {
  const {fileTo, title} = dataConvert;
  let res = constants.NO_ERROR;
  // add S3 download logic later
  if(url){
    const dest = path.join(path.dirname(url), title);
    console.log('fileto:::', fileTo, dest)
    fs.copyFileSync(fileTo, dest);
  }
  return res;
}

function* ExecuteTask(ctx, execObj) {
  var resData;
  var url;
  var tempDirs;
  var getTaskTime = new Date();
  let { cmd, fileTo} = execObj;
  // var cmd = task.getCmd();
  var dataConvert = new TaskQueueDataConvert(ctx, execObj);
  var error = constants.NO_ERROR;
  tempDirs = getTempDir();
  // let fileTo = task.getToFile();
  dataConvert.fileTo = fileTo ? path.join(tempDirs.result, fileTo) : '';
  let builderParams = cmd.getBuilderParams();
  let authorProps = {lastModifiedBy: null, modified: null};
  let isInJwtToken = cmd.getWithAuthorization();
  // error = yield* isUselessConvertion(ctx, task, cmd);
  // if (constants.NO_ERROR !== error) {
  //   ;
  // } else if (cmd.getUrl()) {
  if (cmd.getUrl()) {
    let format = cmd.getFormat();
    dataConvert.fileFrom = path.join(tempDirs.source, dataConvert.key + '.' + format);
    if (checkPathTraversal(ctx, dataConvert.key, tempDirs.source, dataConvert.fileFrom)) {
      url = cmd.getUrl();
      let withAuthorization = cmd.getWithAuthorization();
      let headers;
      let fileSize;
      // let wopiParams = cmd.getWopiParams();
      // if (wopiParams) {
      //   withAuthorization = false;
      //   isInJwtToken = true;
      //   let fileInfo = wopiParams.commonInfo?.fileInfo;
      //   fileSize = fileInfo?.Size;
      //   ({url, headers} = yield wopiUtils.getWopiFileUrl(ctx, fileInfo, wopiParams.userAuth));
      // }
      if (undefined === fileSize || fileSize > 0) {
        // error = yield* downloadFile(ctx, url, dataConvert.fileFrom, withAuthorization, isInJwtToken, headers);
        error = downloadFile(ctx, url, dataConvert, withAuthorization, isInJwtToken, headers);
        console.log('after download', error, constants.NO_ERROR);
      }
      // if (constants.NO_ERROR === error) {
      //   yield* replaceEmptyFile(ctx, dataConvert.fileFrom, format, cmd.getLCID());
      // }
    } else {
      error = constants.CONVERT_PARAMS;
    }
  } else if (builderParams) {
    //in cause script in POST body
    // yield* downloadFileFromStorage(ctx, cmd.getDocId(), tempDirs.source);
    console.debug('downloadFileFromStorage complete');
    // yield* downloadFile(ctx, url, dataConvert.fileFrom, withAuthorization, isInJwtToken, headers);
    downloadFile(ctx, url, dataConvert);
    // let list = yield utils.listObjects(tempDirs.source, false);
    // if (list.length > 0) {
    //   dataConvert.fileFrom = list[0];
    // }
  } else {
    error = constants.UNKNOWN;
  }
  let childRes = null;
  let isTimeout = false;
  if (constants.NO_ERROR === error) {
    ({childRes, isTimeout} = yield* spawnProcess(ctx, builderParams, tempDirs, dataConvert, authorProps, getTaskTime, {}, isInJwtToken));
  }
  uploadFile(ctx, url, dataConvert);
  if (tempDirs) {
      fs.rmSync(tempDirs.temp, { recursive: true, force: true });
      ctx.logger.debug('deleteFolderRecursive');
  }
  // resData = yield* postProcess(ctx, cmd, dataConvert, tempDirs, childRes, error, isTimeout);
  // return resData;
  return {childRes, isTimeout};
}

function* spawnProcess(ctx, builderParams, tempDirs, dataConvert, authorProps, getTaskTime, task, isInJwtToken) {
  const tenX2tPath = ctx.getCfg('FileConverter.converter.x2tPath', cfgX2tPath);
  const tenDocbuilderPath = ctx.getCfg('FileConverter.converter.docbuilderPath', cfgDocbuilderPath);
  const tenArgs = ctx.getCfg('FileConverter.converter.args', cfgArgs);
  console.log('spawnProcess::', { tenX2tPath, tenDocbuilderPath, builderParams, tempDirs, dataConvert, authorProps, getTaskTime, task, isInJwtToken });
  let childRes, isTimeout = false;
  let childArgs;
  if (tenArgs.length > 0) {
    childArgs = tenArgs.trim().replace(/  +/g, ' ').split(' ');
  } else {
    childArgs = [];
  }
  let processPath;
  if (!builderParams) {
    processPath = tenX2tPath;
    let paramsFile = path.join(tempDirs.temp, 'params.xml');
    dataConvert.serialize(ctx, paramsFile);
    childArgs.push(paramsFile);
    let hiddenXml = yield dataConvert.serializeHidden(ctx);
    if (hiddenXml) {
      childArgs.push(hiddenXml);
    }
  } else {
    fs.mkdirSync(path.join(tempDirs.result, 'output'));
    processPath = tenDocbuilderPath;
    childArgs.push('--check-fonts=0');
    childArgs.push('--save-use-only-names=' + tempDirs.result + '/output');
    if (builderParams.argument) {
      childArgs.push(`--argument=${JSON.stringify(builderParams.argument)}`);
    }
    childArgs.push('--options=' + dataConvert.serializeOptions(ctx, isInJwtToken));
    childArgs.push(dataConvert.fileFrom);
  }
  let timeoutId;
  try {
    const tenSpawnOptions = ctx.getCfg('FileConverter.converter.spawnOptions', cfgSpawnOptions);
    //copy to avoid modification of global cfgSpawnOptions
    let spawnOptions = Object.assign({}, tenSpawnOptions);;
    spawnOptions.env = Object.assign({}, process.env, spawnOptions.env);
    if (authorProps.lastModifiedBy && authorProps.modified) {
      spawnOptions.env['LAST_MODIFIED_BY'] = authorProps.lastModifiedBy;
      spawnOptions.env['MODIFIED'] = authorProps.modified;
    }
    console.log('spawnAsyncPromise::', { processPath, childArgs, spawnOptions});
    let spawnAsyncPromise = spawnAsync(processPath, childArgs, spawnOptions);
    childRes = spawnAsyncPromise.child;
    // let waitMS = Math.max(0, task.getVisibilityTimeout() * 1000 - (new Date().getTime() - getTaskTime.getTime()));
    // timeoutId = setTimeout(function() {
    //   isTimeout = true;
    //   timeoutId = undefined;
    //   //close stdio streams to enable emit 'close' event even if HtmlFileInternal is hung-up
    //   childRes.stdin.end();
    //   childRes.stdout.destroy();
    //   childRes.stderr.destroy();
    //   childRes.kill();
    // }, waitMS);
    childRes = yield spawnAsyncPromise;
  } catch (err) {
    if (null === err.status) {
      console.error('error spawnAsync %s', err.stack);
    } else {
      console.debug('error spawnAsync %s', err.stack);
    }
    childRes = err;
  }
  if (undefined !== timeoutId) {
    clearTimeout(timeoutId);
  }
  return {childRes: childRes, isTimeout: isTimeout};
}
// function* postProcess(ctx, cmd, dataConvert, tempDirs, childRes, error, isTimeout) {
//   var exitCode = 0;
//   var exitSignal = null;
//   if(childRes) {
//     exitCode = childRes.status;
//     exitSignal = childRes.signal;
//   }
//   //CONVERT_CELLLIMITS is not an error, but an indicator that data was lost during opening (can be displayed as an error)
//   if ((0 !== exitCode && constants.CONVERT_CELLLIMITS !== -exitCode) || null !== exitSignal) {
//     if (-1 !== exitCodesReturn.indexOf(-exitCode)) {
//       error = -exitCode;
//     } else if(isTimeout) {
//       error = constants.CONVERT_TIMEOUT;
//     } else {
//       error = constants.CONVERT;
//     }
//     if (-1 !== exitCodesMinorError.indexOf(error)) {
//       console.error(ctx, childRes, true);
//       console.debug('ExitCode (code=%d;signal=%s;error:%d)', exitCode, exitSignal, error);
//     } else {
//       console.error(ctx, childRes, false);
//       console.error('ExitCode (code=%d;signal=%s;error:%d)', exitCode, exitSignal, error);
//     }
//   } else {
//     console.error(ctx, childRes, true);
//     console.debug('ExitCode (code=%d;signal=%s;error:%d)', exitCode, exitSignal, error);
//   }
//   if (-1 !== exitCodesUpload.indexOf(error)) {
//     if (-1 !== exitCodesCopyOrigin.indexOf(error)) {
//       let originPath = path.join(path.dirname(dataConvert.fileTo), "origin" + path.extname(dataConvert.fileFrom));
//       if (!fs.existsSync(dataConvert.fileTo)) {
//         fs.copyFileSync(dataConvert.fileFrom, originPath);
//         console.debug('copyOrigin complete');
//       }
//     }
//     console.debug('processUploadToStorage complete');
//   }
//   cmd.setStatusInfo(error);
//   var existFile = false;
//   try {
//     existFile = fs.lstatSync(dataConvert.fileTo).isFile();
//   } catch (err) {
//     existFile = false;
//   }
//   if (!existFile) {
//     //todo review. the stub in the case of AVS_OFFICESTUDIO_FILE_OTHER_OOXML x2t changes the file extension.
//     var fileToBasename = path.basename(dataConvert.fileTo, path.extname(dataConvert.fileTo));
//     var fileToDir = path.dirname(dataConvert.fileTo);
//     var files = fs.readdirSync(fileToDir);
//     for (var i = 0; i < files.length; ++i) {
//       var fileCur = files[i];
//       if (0 == fileCur.indexOf(fileToBasename)) {
//         dataConvert.fileTo = path.join(fileToDir, fileCur);
//         break;
//       }
//     }
//   }
//   cmd.setOutputPath(path.basename(dataConvert.fileTo));
//   if(!cmd.getTitle()){
//     cmd.setTitle(cmd.getOutputPath());
//   }

//   var queueData = new commonDefines.TaskQueueData();
//   queueData.setCtx(ctx);
//   queueData.setCmd(cmd);
//   console.debug('output (data=%j)', queueData);
//   return queueData;
// }

module.exports = router;