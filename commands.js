/******/ (function() { // webpackBootstrap
var __webpack_exports__ = {};
/*!*************************!*\
  !*** ./src/commands.js ***!
  \*************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function ownKeys(e, r) { var t = Object.keys(e); if (Object.getOwnPropertySymbols) { var o = Object.getOwnPropertySymbols(e); r && (o = o.filter(function (r) { return Object.getOwnPropertyDescriptor(e, r).enumerable; })), t.push.apply(t, o); } return t; }
function _objectSpread(e) { for (var r = 1; r < arguments.length; r++) { var t = null != arguments[r] ? arguments[r] : {}; r % 2 ? ownKeys(Object(t), !0).forEach(function (r) { _defineProperty(e, r, t[r]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys(Object(t)).forEach(function (r) { Object.defineProperty(e, r, Object.getOwnPropertyDescriptor(t, r)); }); } return e; }
function _defineProperty(e, r, t) { return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: !0, configurable: !0, writable: !0 }) : e[r] = t, e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return e; }; var t, e = {}, r = Object.prototype, n = r.hasOwnProperty, o = Object.defineProperty || function (t, e, r) { t[e] = r.value; }, i = "function" == typeof Symbol ? Symbol : {}, a = i.iterator || "@@iterator", c = i.asyncIterator || "@@asyncIterator", u = i.toStringTag || "@@toStringTag"; function define(t, e, r) { return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e]; } try { define({}, ""); } catch (t) { define = function define(t, e, r) { return t[e] = r; }; } function wrap(t, e, r, n) { var i = e && e.prototype instanceof Generator ? e : Generator, a = Object.create(i.prototype), c = new Context(n || []); return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a; } function tryCatch(t, e, r) { try { return { type: "normal", arg: t.call(e, r) }; } catch (t) { return { type: "throw", arg: t }; } } e.wrap = wrap; var h = "suspendedStart", l = "suspendedYield", f = "executing", s = "completed", y = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var p = {}; define(p, a, function () { return this; }); var d = Object.getPrototypeOf, v = d && d(d(values([]))); v && v !== r && n.call(v, a) && (p = v); var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p); function defineIteratorMethods(t) { ["next", "throw", "return"].forEach(function (e) { define(t, e, function (t) { return this._invoke(e, t); }); }); } function AsyncIterator(t, e) { function invoke(r, o, i, a) { var c = tryCatch(t[r], t, o); if ("throw" !== c.type) { var u = c.arg, h = u.value; return h && "object" == _typeof(h) && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) { invoke("next", t, i, a); }, function (t) { invoke("throw", t, i, a); }) : e.resolve(h).then(function (t) { u.value = t, i(u); }, function (t) { return invoke("throw", t, i, a); }); } a(c.arg); } var r; o(this, "_invoke", { value: function value(t, n) { function callInvokeWithMethodAndArg() { return new e(function (e, r) { invoke(t, n, e, r); }); } return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(e, r, n) { var o = h; return function (i, a) { if (o === f) throw Error("Generator is already running"); if (o === s) { if ("throw" === i) throw a; return { value: t, done: !0 }; } for (n.method = i, n.arg = a;;) { var c = n.delegate; if (c) { var u = maybeInvokeDelegate(c, n); if (u) { if (u === y) continue; return u; } } if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) { if (o === h) throw o = s, n.arg; n.dispatchException(n.arg); } else "return" === n.method && n.abrupt("return", n.arg); o = f; var p = tryCatch(e, r, n); if ("normal" === p.type) { if (o = n.done ? s : l, p.arg === y) continue; return { value: p.arg, done: n.done }; } "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg); } }; } function maybeInvokeDelegate(e, r) { var n = r.method, o = e.iterator[n]; if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y; var i = tryCatch(o, e.iterator, r.arg); if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y; var a = i.arg; return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y); } function pushTryEntry(t) { var e = { tryLoc: t[0] }; 1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e); } function resetTryEntry(t) { var e = t.completion || {}; e.type = "normal", delete e.arg, t.completion = e; } function Context(t) { this.tryEntries = [{ tryLoc: "root" }], t.forEach(pushTryEntry, this), this.reset(!0); } function values(e) { if (e || "" === e) { var r = e[a]; if (r) return r.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) { var o = -1, i = function next() { for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next; return next.value = t, next.done = !0, next; }; return i.next = i; } } throw new TypeError(_typeof(e) + " is not iterable"); } return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), o(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) { var e = "function" == typeof t && t.constructor; return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name)); }, e.mark = function (t) { return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t; }, e.awrap = function (t) { return { __await: t }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () { return this; }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) { void 0 === i && (i = Promise); var a = new AsyncIterator(wrap(t, r, n, o), i); return e.isGeneratorFunction(r) ? a : a.next().then(function (t) { return t.done ? t.value : a.next(); }); }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () { return this; }), define(g, "toString", function () { return "[object Generator]"; }), e.keys = function (t) { var e = Object(t), r = []; for (var n in e) r.push(n); return r.reverse(), function next() { for (; r.length;) { var t = r.pop(); if (t in e) return next.value = t, next.done = !1, next; } return next.done = !0, next; }; }, e.values = values, Context.prototype = { constructor: Context, reset: function reset(e) { if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t); }, stop: function stop() { this.done = !0; var t = this.tryEntries[0].completion; if ("throw" === t.type) throw t.arg; return this.rval; }, dispatchException: function dispatchException(e) { if (this.done) throw e; var r = this; function handle(n, o) { return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o; } for (var o = this.tryEntries.length - 1; o >= 0; --o) { var i = this.tryEntries[o], a = i.completion; if ("root" === i.tryLoc) return handle("end"); if (i.tryLoc <= this.prev) { var c = n.call(i, "catchLoc"), u = n.call(i, "finallyLoc"); if (c && u) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } else if (c) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); } else { if (!u) throw Error("try statement without catch or finally"); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } } } }, abrupt: function abrupt(t, e) { for (var r = this.tryEntries.length - 1; r >= 0; --r) { var o = this.tryEntries[r]; if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) { var i = o; break; } } i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null); var a = i ? i.completion : {}; return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a); }, complete: function complete(t, e) { if ("throw" === t.type) throw t.arg; return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y; }, finish: function finish(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y; } }, catch: function _catch(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.tryLoc === t) { var n = r.completion; if ("throw" === n.type) { var o = n.arg; resetTryEntry(r); } return o; } } throw Error("illegal catch attempt"); }, delegateYield: function delegateYield(e, r, n) { return this.delegate = { iterator: values(e), resultName: r, nextLoc: n }, "next" === this.method && (this.arg = t), y; } }, e; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var config = {
  clientId: "6d7e781e-9cf5-48ff-8c05-b697ca1a90e3",
  redirectUri: "https://nyusen.github.io/Toggle-Save-For-Email/auth-callback.html",
  authEndpoint: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
  tokenEndpoint: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
  scopes: "openid profile email Mail.Read Mail.Send"
};
var accessToken = null;
var idToken = null;

// Make functions globally available
window.validateBody = validateBody;
window.saveForTraining = saveForTraining;
window.setSendOnly = setSendOnly;
Office.onReady(/*#__PURE__*/function () {
  var _ref = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee(info) {
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          if (info.host) {
            // Check if we have tokens in session storage
            accessToken = sessionStorage.getItem('accessToken');
            idToken = sessionStorage.getItem('idToken');
            if (accessToken) {
              updateUI(true);
            } else {
              updateUI(false);
            }

            // Add click handler for sign-in button
            document.getElementById('signInButton').onclick = handleSignIn;
          }
        case 1:
        case "end":
          return _context.stop();
      }
    }, _callee);
  }));
  return function (_x) {
    return _ref.apply(this, arguments);
  };
}());

// Function to handle sign-in button click
function handleSignIn() {
  return new Promise(function (resolve, reject) {
    // Generate random state and PKCE verifier
    var state = generateRandomString(32);
    var codeVerifier = generateRandomString(64);
    var nonce = generateRandomString(16);

    // Store state and code verifier in parent window's session storage
    window.sessionStorage.setItem('authState', state);
    window.sessionStorage.setItem('codeVerifier', codeVerifier);

    // Generate code challenge
    generateCodeChallenge(codeVerifier).then(function (codeChallenge) {
      // Build the URL to our local sign-in start page
      var startUrl = new URL('https://nyusen.github.io/Toggle-Save-For-Email/sign-in-start.html');
      startUrl.searchParams.append('state', state);
      startUrl.searchParams.append('nonce', nonce);
      startUrl.searchParams.append('code_challenge', codeChallenge);
      startUrl.searchParams.append('code_challenge_method', 'S256');
      startUrl.searchParams.append('code_verifier', codeVerifier);

      // Add delay before opening the auth dialog
      setTimeout(function () {
        // Open our local page in a dialog, which will then redirect to Microsoft
        Office.context.ui.displayDialogAsync(startUrl.toString(), {
          height: 60,
          width: 30
        }, function (result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error(result.error.message);
            reject(new Error(result.error.message));
            return;
          }
          var dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            try {
              var message = JSON.parse(arg.message);
              if (message.type === 'token') {
                // Store both tokens
                accessToken = message.token;
                idToken = message.idToken;
                sessionStorage.setItem('accessToken', message.token);
                sessionStorage.setItem('idToken', message.idToken);
                sessionStorage.setItem('tokenType', message.tokenType);
                sessionStorage.setItem('expiresIn', message.expiresIn);
                dialog.close();
                resolve(message.token);
              } else if (message.type === 'error') {
                dialog.close();
                reject(new Error(message.error));
              }
            } catch (e) {
              console.error('Error parsing message:', e);
              dialog.close();
              reject(e);
            }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
            dialog.close();
            reject(new Error('Dialog closed'));
          });
        });
      }, 1000); // Wait 1 second to ensure previous dialog is fully closed
    }).catch(function (error) {
      reject(error);
    });
  });
}

// Function to generate random string for PKCE and state
function generateRandomString(length) {
  var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
  var text = '';
  for (var i = 0; i < length; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

// Function to generate code challenge from verifier
function generateCodeChallenge(_x2) {
  return _generateCodeChallenge.apply(this, arguments);
} // Base64URL encoding function
function _generateCodeChallenge() {
  _generateCodeChallenge = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee2(codeVerifier) {
    var encoder, data, digest;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          encoder = new TextEncoder();
          data = encoder.encode(codeVerifier);
          _context2.next = 4;
          return window.crypto.subtle.digest('SHA-256', data);
        case 4:
          digest = _context2.sent;
          return _context2.abrupt("return", base64URLEncode(digest));
        case 6:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return _generateCodeChallenge.apply(this, arguments);
}
function base64URLEncode(buffer) {
  return btoa(String.fromCharCode.apply(String, _toConsumableArray(new Uint8Array(buffer)))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Function to update UI based on sign-in state
function updateUI(isSignedIn) {
  var signInSection = document.getElementById('signInSection');
  var toggleSection = document.getElementById('toggleSection');
  var statusMessage = document.getElementById('statusMessage');
  var signInButton = document.getElementById('signInButton');
  if (isSignedIn) {
    signInButton.style.display = 'none';
    toggleSection.style.display = 'block';
    statusMessage.textContent = 'Signed in successfully';
  } else {
    signInButton.style.display = 'block';
    toggleSection.style.display = 'none';
    statusMessage.textContent = 'Please sign in to use the add-in';
  }
}

// Function to make authenticated requests to your server
function makeAuthenticatedRequest(_x3) {
  return _makeAuthenticatedRequest.apply(this, arguments);
} // Function to set send only option
function _makeAuthenticatedRequest() {
  _makeAuthenticatedRequest = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee3(url) {
    var options,
      headers,
      response,
      _args3 = arguments;
    return _regeneratorRuntime().wrap(function _callee3$(_context3) {
      while (1) switch (_context3.prev = _context3.next) {
        case 0:
          options = _args3.length > 1 && _args3[1] !== undefined ? _args3[1] : {};
          if (idToken) {
            _context3.next = 3;
            break;
          }
          throw new Error('Not authenticated');
        case 3:
          headers = _objectSpread({
            'Authorization': "Bearer ".concat(idToken),
            'Content-Type': 'application/json'
          }, options.headers);
          _context3.next = 6;
          return fetch(url, _objectSpread(_objectSpread({}, options), {}, {
            headers: headers
          }));
        case 6:
          response = _context3.sent;
          if (response.ok) {
            _context3.next = 14;
            break;
          }
          if (!(response.status === 401)) {
            _context3.next = 13;
            break;
          }
          // Token might be expired, clear it
          accessToken = null;
          sessionStorage.removeItem('accessToken');
          updateUI(false);
          throw new Error('Authentication expired. Please sign in again.');
        case 13:
          throw new Error("Request failed: ".concat(response.statusText));
        case 14:
          return _context3.abrupt("return", response.json());
        case 15:
        case "end":
          return _context3.stop();
      }
    }, _callee3);
  }));
  return _makeAuthenticatedRequest.apply(this, arguments);
}
function setSendOnly(event) {
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
    var props = result.value;
    props.set("saveForTraining", false);
    props.saveAsync(function () {
      event.completed();
    });
  });
}

// Entry point for email send validation
function validateBody(_x4) {
  return _validateBody.apply(this, arguments);
}
function _validateBody() {
  _validateBody = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee6(event) {
    return _regeneratorRuntime().wrap(function _callee6$(_context6) {
      while (1) switch (_context6.prev = _context6.next) {
        case 0:
          try {
            // Load custom properties to check save setting
            Office.context.mailbox.item.loadCustomPropertiesAsync(/*#__PURE__*/function () {
              var _ref2 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee5(result) {
                var props, shouldSave;
                return _regeneratorRuntime().wrap(function _callee5$(_context5) {
                  while (1) switch (_context5.prev = _context5.next) {
                    case 0:
                      props = result.value;
                      shouldSave = props.get("saveForTraining");
                      if (shouldSave === undefined) {
                        event.completed({
                          allowEvent: true
                        });
                      }
                      if (shouldSave) {
                        _context5.next = 6;
                        break;
                      }
                      // If "Send Only" is selected, just send the email
                      event.completed({
                        allowEvent: true
                      });
                      return _context5.abrupt("return");
                    case 6:
                      if (accessToken) {
                        _context5.next = 10;
                        break;
                      }
                      // Show sign-in dialog
                      Office.context.ui.displayDialogAsync('https://nyusen.github.io/Toggle-Save-For-Email/signin-dialog.html', {
                        height: 40,
                        width: 30,
                        displayInIframe: true
                      }, function (signInResult) {
                        if (signInResult.status === Office.AsyncResultStatus.Failed) {
                          console.error(signInResult.error.message);
                          event.completed({
                            allowEvent: false
                          });
                          return;
                        }
                        var signInDialog = signInResult.value;
                        signInDialog.addEventHandler(Office.EventType.DialogMessageReceived, /*#__PURE__*/function () {
                          var _ref3 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee4(signInArg) {
                            return _regeneratorRuntime().wrap(function _callee4$(_context4) {
                              while (1) switch (_context4.prev = _context4.next) {
                                case 0:
                                  if (signInArg.message === 'signin') {
                                    // User wants to sign in
                                    signInDialog.close();
                                    handleSignIn().then(function () {
                                      // User is signed in, proceed with saving
                                      saveForTraining(event);
                                    }).catch(function (error) {
                                      console.error('Error signing in:', error);
                                      event.completed({
                                        allowEvent: false
                                      });
                                    });
                                  } else if (signInArg.message === 'cancel') {
                                    // User doesn't want to sign in
                                    signInDialog.close();
                                    event.completed({
                                      allowEvent: false
                                    });
                                  }
                                case 1:
                                case "end":
                                  return _context4.stop();
                              }
                            }, _callee4);
                          }));
                          return function (_x7) {
                            return _ref3.apply(this, arguments);
                          };
                        }());
                      });
                      _context5.next = 12;
                      break;
                    case 10:
                      _context5.next = 12;
                      return saveForTraining(event);
                    case 12:
                    case "end":
                      return _context5.stop();
                  }
                }, _callee5);
              }));
              return function (_x6) {
                return _ref2.apply(this, arguments);
              };
            }());
          } catch (error) {
            console.error('Error in validateBody:', error);
            event.completed({
              allowEvent: true,
              errorMessage: "Failed to process email, but allowing send."
            });
          }
        case 1:
        case "end":
          return _context6.stop();
      }
    }, _callee6);
  }));
  return _validateBody.apply(this, arguments);
}
function saveForTraining(_x5) {
  return _saveForTraining.apply(this, arguments);
}
function _saveForTraining() {
  _saveForTraining = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee8(event) {
    var item;
    return _regeneratorRuntime().wrap(function _callee8$(_context8) {
      while (1) switch (_context8.prev = _context8.next) {
        case 0:
          item = Office.context.mailbox.item;
          item.loadCustomPropertiesAsync(/*#__PURE__*/function () {
            var _ref4 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee7(result) {
              var customProperties, tags, _yield$Promise$all, _yield$Promise$all2, subject, body, sender, recipients, emailData, errorMessage;
              return _regeneratorRuntime().wrap(function _callee7$(_context7) {
                while (1) switch (_context7.prev = _context7.next) {
                  case 0:
                    _context7.prev = 0;
                    customProperties = result.value;
                    tags = JSON.parse(customProperties.get('selectedTags')) || []; // Get email data using promises
                    _context7.next = 5;
                    return Promise.all([new Promise(function (resolve) {
                      return item.subject.getAsync(resolve);
                    }), new Promise(function (resolve) {
                      return item.body.getAsync('text', resolve);
                    }), new Promise(function (resolve) {
                      return item.from.getAsync(resolve);
                    }), new Promise(function (resolve) {
                      return item.to.getAsync(resolve);
                    })]);
                  case 5:
                    _yield$Promise$all = _context7.sent;
                    _yield$Promise$all2 = _slicedToArray(_yield$Promise$all, 4);
                    subject = _yield$Promise$all2[0];
                    body = _yield$Promise$all2[1];
                    sender = _yield$Promise$all2[2];
                    recipients = _yield$Promise$all2[3];
                    emailData = {
                      subject: subject.value,
                      body: body.value,
                      sender: sender.value.emailAddress,
                      recipients: recipients.value.map(function (r) {
                        return r.emailAddress;
                      }),
                      tags: tags.map(function (tag) {
                        return tag.id;
                      }),
                      timestamp: new Date().toISOString()
                    }; // Make authenticated request to your server
                    _context7.next = 14;
                    return makeAuthenticatedRequest('https://ml-inf-svc-dev.eventellect.com/corpus-collector/api/save-email', {
                      method: 'POST',
                      headers: {
                        'Content-Type': 'application/json'
                      },
                      body: JSON.stringify(emailData)
                    });
                  case 14:
                    event.completed({
                      allowEvent: true
                    });
                    _context7.next = 22;
                    break;
                  case 17:
                    _context7.prev = 17;
                    _context7.t0 = _context7["catch"](0);
                    console.error('Error:', _context7.t0);
                    // Show retry dialog with error message
                    errorMessage = encodeURIComponent(_context7.t0.message || 'An unknown error occurred');
                    Office.context.ui.displayDialogAsync("https://nyusen.github.io/Toggle-Save-For-Email/retry-dialog.html?error=".concat(errorMessage), {
                      height: 30,
                      width: 20,
                      displayInIframe: true
                    }, function (asyncResult) {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error(asyncResult.error.message);
                        event.completed({
                          allowEvent: false
                        });
                        return;
                      }
                      var dialog = asyncResult.value;
                      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                        dialog.close();
                        if (arg.message === 'retry') {
                          saveForTraining(event);
                        } else {
                          event.completed({
                            allowEvent: true
                          });
                        }
                      });
                    });
                  case 22:
                  case "end":
                    return _context7.stop();
                }
              }, _callee7, null, [[0, 17]]);
            }));
            return function (_x8) {
              return _ref4.apply(this, arguments);
            };
          }());
        case 2:
        case "end":
          return _context8.stop();
      }
    }, _callee8);
  }));
  return _saveForTraining.apply(this, arguments);
}
window.commands = __webpack_exports__;
/******/ })()
;
//# sourceMappingURL=commands.js.map