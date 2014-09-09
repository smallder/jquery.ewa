/**
 * Excel Services ECMAScript (Ewa) jQuery integration Plugin
 * To see more Ewa API you should go to http://msdn.microsoft.com/en-us/library/ee589018.aspx
 * 
 * Version: 1.0.0
 * Author: http://smallder.ru
 * Source: https://github.com/smallder/jquery.ewa
 * Copyright (c) 2014 smallder; GNU GPL v2
 */

'use strict';

(function($) {
	$.fn.ewa = function(config) {
		/* Check direct method call */
		if (typeof config == 'string') {
			return this.triggerHandler(config, arguments[1]);
		}
		
		/**
		 * Plugin config
		 * @var object
		 */
		var config = $.extend(true, {
			token: '',
			ewa: {
				uiOptions: {
					showGridlines: true,
					showRowColumnHeaders: true,
					showParametersTaskPane: true
				},
				interactivityOptions: {
					allowTypingAndFormulaEntry: true,
					allowParameterModification: false,
					allowSorting: false,
					allowFiltering: false,
					allowPivotTableInteractivity: false
				}
			},
			load: null
		}, config);
		
		/**
		 * Plugin handled dom elements
		 * @var jQuery
		 */
		var elements = this;
		
		/**
		 * Flexible getRangeA1Async() adapter - works with or without sheet name
		 * @param Ewa.EwaControl ewa Ewa control
		 * @param object addr Range address
		 * @param function callback Async callback
		 * @return void
		 */
		var getRangeA1AsyncFlex = function(ewa, addr, callback) {
			if (addr.split('!').length > 1) {
				ewa.getActiveWorkbook().getRangeA1Async(addr, callback)
			} else {
				ewa.getActiveWorkbook().getActiveSheet().getRangeA1Async(addr, callback)
			}
		};
		
		/**
		 * Ewa sheet loader
		 * @return void
		 */
		var loadEwa = function() {
			elements.each(function() {
				var element = $(this);
				var ewa = null;
				
				if (!element.attr('id')) {
					element.attr('id', 'ewa' + Math.floor(999999999 * Math.random()));
				}
				
				Ewa.EwaControl.loadEwaAsync(
					config.token || element.data('token'),
					element.attr('id'),
					config.ewa,
					function(asyncResult) {
						if (asyncResult.getSucceeded()) {
							ewa = asyncResult.getEwaControl();
						} else {
							throw 'Async operation failed!';
						}
						
						if (config.load) {
							config.load.call(element, arguments);
						}
					}
				);
				
				/**
				 * Return ewa control
				 * @return Ewa.EwaControl
				 */
				element.on('getEwa', function(event) {
					return ewa;
				});
				
				/**
				 * Return col index by A1-style code
				 * @param jQuery.Event event Ewa control
				 * @param string col Col code (A, Sheet1!A, AC, etc.)
				 * @return integer
				 */
				element.on('getColIndexFromA', function(event, col) {
					var startCode = 'A'.charCodeAt(0);
					var finishCode = 'Z'.charCodeAt(0);
					var module = finishCode - startCode;
					
					col = col.toUpperCase();
					var index = 0;
					for (var i = col.length; i > 0; i--) {
						var val = Math.pow(module, col.length - i) * (1 + col.charCodeAt(i - 1) - startCode);
						index += val;
					}
					
					return index - 1;
				});
				
				/**
				 * Return sheet range values
				 * @param jQuery.Event event Ewa control
				 * @param string range Range (A1:C10, Sheet1!A1:C10, etc.)
				 * @param function callback Async callback for handle result
				 * @return void
				 */
				element.on('getRangeValues', function(event, range, callback) {
					getRangeA1AsyncFlex(ewa, range, function(asyncResult) {
						var range = asyncResult.getReturnValue();
						range.getValuesAsync(Ewa.ValuesFormat.Unformatted, function(asyncResult) {
							callback.call(element, asyncResult.getReturnValue(), range, asyncResult);
						}, asyncResult.getUserContext());
					});
				});
				
				/**
				 * Return cell value
				 * @param jQuery.Event event Ewa control
				 * @param string addr Cell A1-address (A1, Sheet1!A1, etc.)
				 * @param function callback Async callback for handle result
				 * @return void
				 */
				element.on('getCellValue', function(event, addr, callback) {
					getRangeA1AsyncFlex(ewa, addr, function(asyncResult) {
						var range = asyncResult.getReturnValue();
						range.getValuesAsync(Ewa.ValuesFormat.Unformatted, function(asyncResult) {
							var value = asyncResult.getReturnValue();
							callback.call(element, value[0][0], asyncResult);
						}, asyncResult.getUserContext());
					});
				});
				
				/**
				 * Set cell value
				 * @param jQuery.Event event Ewa control
				 * @param string addr Cell A1-address (A1, Sheet1!A1, etc.)
				 * @param string value Cell value
				 * @param function|null callback Async callback for handle result
				 * @return void
				 */
				element.on('setCellValue', function(event, addr, value, callback) {
					getRangeA1AsyncFlex(ewa, addr, function(asyncResult) {
						var range = asyncResult.getReturnValue();
						range.setValuesAsync([[value]], function(asyncResult) {
							callback && callback.apply(element, arguments);
						}, asyncResult.getUserContext());
					});
				});
			});
		};
		
		/* Check if Ewa namespace is exists */
		if (typeof Ewa == 'undefined') {
			var protocol = document.location.protocol == 'https:' ? 'https:' : 'http:';
			$.getScript(protocol + '//r.office.microsoft.com/r/rlidExcelWLJS?v=1&kip=1', loadEwa).fail(function(jqxhr, settings, exception) {
				throw 'Can\'t load Ewa API with error #' + jqxhr.status + ' "' + jqxhr.statusText + '"';
			});
		} else {
			loadEwa();
		}
		
		return this;
	};
})(jQuery);