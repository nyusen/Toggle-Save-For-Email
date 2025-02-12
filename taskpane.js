/******/ (function() { // webpackBootstrap
var __webpack_exports__ = {};
/*!*****************************!*\
  !*** ./src/tag-taskpane.js ***!
  \*****************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
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
// Global variables for authentication
var accessToken = null;
var idToken = null;

// Store tags globally so we can filter them
var availableTags = [];

// Store the authenticated HTML content
var authenticatedHTML = "\n    <div class=\"tag-container\">\n        <div class=\"section\">\n            <h2 class=\"section-title\">All Tags</h2>\n            <div class=\"search-container\">\n                <input type=\"text\" \n                       id=\"tagSearch\" \n                       class=\"search-input\" \n                       placeholder=\"Search tags...\"\n                       oninput=\"filterTags(this.value)\">\n            </div>\n            <div id=\"tagList\" class=\"tag-list\">\n                <!-- Tags will be populated here -->\n            </div>\n        </div>\n        \n        <div class=\"section\">\n            <h2 class=\"section-title\">Add a Tag</h2>\n            <div class=\"add-tag-container\">\n                <input type=\"text\" \n                       id=\"customTagInput\" \n                       class=\"custom-tag-input\" \n                       placeholder=\"Enter a custom tag\"\n                       oninput=\"toggleAddButton()\">\n                <button id=\"addTagButton\" \n                        class=\"add-tag-button\" \n                        onclick=\"addCustomTag()\" \n                        disabled>Add Tag</button>\n                <div id=\"tagError\" class=\"tag-error\">This tag already exists</div>\n            </div>\n        </div>\n\n        <div class=\"section\">\n            <h2 class=\"section-title\">Selected Tags</h2>\n            <div id=\"selectedTags\" class=\"selected-tags\">\n                <!-- Selected tags will appear here -->\n            </div>\n        </div>\n    </div>\n";
Office.onReady(function () {
  // Check if we have tokens in session storage
  accessToken = sessionStorage.getItem('accessToken');
  idToken = sessionStorage.getItem('idToken');

  // Set saveForTraining to true when taskpane loads
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var props = result.value;
      props.set("saveForTraining", true);
      props.saveAsync(function () {
        initializeUI();
      });
    }
  });
});
function initializeUI() {
  if (idToken) {
    // If authenticated, show the tag interface and load tags
    document.body.innerHTML = authenticatedHTML;
    loadTags();
  } else {
    // If not signed in, show sign-in dialog
    document.body.innerHTML = '<div class="error-message">Please sign in to manage tags.</div>';
    showSignInDialog();
  }
}
function showSignInDialog() {
  Office.context.ui.displayDialogAsync('https://nyusen.github.io/Toggle-Save-For-Email/signin-dialog.html', {
    height: 40,
    width: 30,
    displayInIframe: true
  }, function (signInResult) {
    if (signInResult.status === Office.AsyncResultStatus.Failed) {
      console.error(signInResult.error.message);
      showError('Failed to open sign-in dialog');
      return;
    }
    var signInDialog = signInResult.value;
    signInDialog.addEventHandler(Office.EventType.DialogMessageReceived, /*#__PURE__*/function () {
      var _ref = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee(signInArg) {
        return _regeneratorRuntime().wrap(function _callee$(_context) {
          while (1) switch (_context.prev = _context.next) {
            case 0:
              if (signInArg.message === 'signin') {
                // User wants to sign in
                signInDialog.close();
                // Add timeout before proceeding
                setTimeout(function () {
                  handleSignIn().then(function () {
                    // User is signed in, restore the authenticated UI and load tags
                    setTimeout(function () {
                      initializeUI();
                    }, 1000);
                  }).catch(function (error) {
                    showError('Failed to sign in');
                  });
                }, 1000);
              } else if (signInArg.message === 'cancel') {
                // User doesn't want to sign in
                signInDialog.close();
                setTimeout(function () {
                  showError('Authentication required to load tags');
                }, 1000);
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
  });
}

// Function to handle sign-in
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

// Function to show error message in the taskpane
function showError(message) {
  var container = document.querySelector('.tag-container');
  container.innerHTML = "\n        <div style=\"color: red; padding: 20px; text-align: center;\">\n            <p>".concat(message, "</p>\n            <button onclick=\"showSignInDialog()\">Sign In</button>\n        </div>\n    ");
}

// Function to fetch tags from the server
function loadTags() {
  return _loadTags.apply(this, arguments);
} // Function to display filtered tags
function _loadTags() {
  _loadTags = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee2() {
    var response, tagList;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          showTagListLoading();
          _context2.prev = 1;
          _context2.next = 4;
          return makeAuthenticatedRequest('https://ml-inf-svc-dev.eventellect.com/corpus-collector/api/tag');
        case 4:
          response = _context2.sent;
          _context2.next = 7;
          return response.json();
        case 7:
          availableTags = _context2.sent;
          displayFilteredTags(availableTags);
          _context2.next = 16;
          break;
        case 11:
          _context2.prev = 11;
          _context2.t0 = _context2["catch"](1);
          console.error('Error loading tags:', _context2.t0);
          tagList = document.getElementById('tagList');
          tagList.innerHTML = '<div class="error-message">Failed to load tags. Please try again later.</div>';
        case 16:
        case "end":
          return _context2.stop();
      }
    }, _callee2, null, [[1, 11]]);
  }));
  return _loadTags.apply(this, arguments);
}
function displayFilteredTags(tags) {
  var tagList = document.getElementById('tagList');
  if (tags.length === 0) {
    tagList.innerHTML = '<div class="no-results">No matching tags found</div>';
    return;
  }

  // Sort tags alphabetically by description, ignoring case
  var sortedTags = _toConsumableArray(tags).sort(function (a, b) {
    return a.description.toLowerCase().localeCompare(b.description.toLowerCase());
  });

  // Get currently selected tags from custom properties
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var props = result.value;
      var selectedTags = props.get('selectedTags') || [];
      if (typeof selectedTags === 'string') {
        selectedTags = JSON.parse(selectedTags);
      }

      // Create tag list with checked state preserved
      tagList.innerHTML = sortedTags.map(function (tag) {
        var isChecked = selectedTags.some(function (selectedTag) {
          return selectedTag.id === tag.id;
        });
        return "\n                    <div class=\"tag-item\">\n                        <input type=\"checkbox\" id=\"tag-".concat(tag.id, "\" value=\"").concat(tag.id, "\" \n                            onchange=\"handleTagSelection(this)\" ").concat(isChecked ? 'checked' : '', ">\n                        <label for=\"tag-").concat(tag.id, "\">").concat(tag.description, "</label>\n                    </div>\n                ");
      }).join('');
    }
  });
}

// Function to filter tags based on search input
function filterTags(searchText) {
  if (!availableTags) return;
  var searchLower = searchText.toLowerCase().trim();
  var filteredTags = searchLower === '' ? availableTags : availableTags.filter(function (tag) {
    return tag.description.toLowerCase().includes(searchLower);
  });
  displayFilteredTags(filteredTags);

  // Restore checked state for selected tags
  var selectedTags = getSelectedTags();
  selectedTags.forEach(function (tag) {
    var checkbox = document.getElementById("tag-".concat(tag.id));
    if (checkbox) {
      checkbox.checked = true;
    }
  });
}

// Function to get currently selected tags
function getSelectedTags() {
  var selectedTagsContainer = document.getElementById('selectedTags');
  return Array.from(selectedTagsContainer.getElementsByClassName('selected-tag')).map(function (tagElement) {
    var tagId = parseInt(tagElement.dataset.tagId);
    return availableTags.find(function (tag) {
      return tag.id === tagId;
    });
  }).filter(function (tag) {
    return tag;
  }); // Remove any undefined entries
}

// Function to make authenticated requests
function makeAuthenticatedRequest(_x2) {
  return _makeAuthenticatedRequest.apply(this, arguments);
} // Utility functions for authentication
function _makeAuthenticatedRequest() {
  _makeAuthenticatedRequest = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee3(url) {
    var options,
      requestOptions,
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
          throw new Error('No id token available');
        case 3:
          requestOptions = _objectSpread(_objectSpread({}, options), {}, {
            headers: _objectSpread(_objectSpread({}, options.headers), {}, {
              'Authorization': "Bearer ".concat(idToken)
            })
          });
          _context3.next = 6;
          return fetch(url, requestOptions);
        case 6:
          response = _context3.sent;
          if (response.status === 401) {
            // Clear tokens if unauthorized
            sessionStorage.removeItem('accessToken');
            sessionStorage.removeItem('idToken');
            accessToken = null;
            idToken = null;
          }
          return _context3.abrupt("return", response);
        case 9:
        case "end":
          return _context3.stop();
      }
    }, _callee3);
  }));
  return _makeAuthenticatedRequest.apply(this, arguments);
}
function generateRandomString(length) {
  var array = new Uint8Array(length);
  crypto.getRandomValues(array);
  return Array.from(array, function (byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}
function generateCodeChallenge(_x3) {
  return _generateCodeChallenge.apply(this, arguments);
}
function _generateCodeChallenge() {
  _generateCodeChallenge = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee4(verifier) {
    var encoder, data, hash;
    return _regeneratorRuntime().wrap(function _callee4$(_context4) {
      while (1) switch (_context4.prev = _context4.next) {
        case 0:
          encoder = new TextEncoder();
          data = encoder.encode(verifier);
          _context4.next = 4;
          return crypto.subtle.digest('SHA-256', data);
        case 4:
          hash = _context4.sent;
          return _context4.abrupt("return", base64URLEncode(hash));
        case 6:
        case "end":
          return _context4.stop();
      }
    }, _callee4);
  }));
  return _generateCodeChallenge.apply(this, arguments);
}
function base64URLEncode(buffer) {
  return btoa(String.fromCharCode.apply(String, _toConsumableArray(new Uint8Array(buffer)))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Function to handle tag selection
function handleTagSelection(checkbox) {
  var tagId = parseInt(checkbox.value);
  var selectedTagsContainer = document.getElementById('selectedTags');
  if (checkbox.checked) {
    // Find the tag object from availableTags
    var selectedTag = availableTags.find(function (tag) {
      return tag.id === tagId;
    });
    if (selectedTag) {
      var tagElement = document.createElement('span');
      tagElement.className = 'selected-tag';
      tagElement.dataset.tagId = selectedTag.id;
      tagElement.dataset.description = selectedTag.description; // Store description for getSelectedTags
      tagElement.innerHTML = "\n                ".concat(selectedTag.description, "\n                <span class=\"remove-tag\" onclick=\"removeTag('").concat(selectedTag.id, "')\">&times;</span>\n            ");
      selectedTagsContainer.appendChild(tagElement);
    }
  } else {
    // Remove the tag
    var existingTag = selectedTagsContainer.querySelector("[data-tag-id=\"".concat(tagId, "\"]"));
    if (existingTag) {
      existingTag.remove();
    }
  }

  // Save selected tags to custom properties
  saveSelectedTags();
}

// Function to remove a tag
function removeTag(tagId) {
  // Uncheck the checkbox
  var checkbox = document.getElementById("tag-".concat(tagId));
  if (checkbox) {
    checkbox.checked = false;
  }

  // Remove the tag from selected tags
  var selectedTagsContainer = document.getElementById('selectedTags');
  var existingTag = selectedTagsContainer.querySelector("[data-tag-id=\"".concat(tagId, "\"]"));
  if (existingTag) {
    existingTag.remove();
  }

  // Save selected tags to custom properties
  saveSelectedTags();
}

// Function to save selected tags to custom properties
function saveSelectedTags() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var props = result.value;
      var selectedTags = getSelectedTags();
      props.set("selectedTags", JSON.stringify(selectedTags));
      props.saveAsync();
    }
  });
}

// Function to load selected tags from custom properties
function loadSelectedTags() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var props = result.value;
      var savedTags = props.get("selectedTags");
      if (savedTags) {
        try {
          var selectedTagIds = JSON.parse(savedTags);
          selectedTagIds.forEach(function (tag) {
            var checkbox = document.getElementById("tag-".concat(tag.id));
            if (checkbox) {
              checkbox.checked = true;
              // Trigger the selection handler to update the UI
              handleTagSelection(checkbox);
            }
          });
        } catch (e) {
          console.error('Error parsing saved tags:', e);
        }
      }
    }
  });
}

// Function to normalize tag text for comparison
function normalizeTagText(text) {
  return text.toLowerCase().replace(/\s+/g, '');
}

// Function to show tag error message
function showTagError() {
  var message = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : 'This tag already exists';
  var errorElement = document.getElementById('tagError');
  errorElement.textContent = message;
  errorElement.classList.add('visible');
  setTimeout(function () {
    errorElement.classList.remove('visible');
  }, 3000);
}

// Function to toggle the Add Tag button state
function toggleAddButton() {
  var input = document.getElementById('customTagInput');
  var button = document.getElementById('addTagButton');
  var errorElement = document.getElementById('tagError');
  if (input.value.trim() !== '') {
    button.classList.add('active');
    button.disabled = false;
  } else {
    button.classList.remove('active');
    button.disabled = true;
  }
}

// Function to set button loading state
function setAddButtonLoading(isLoading) {
  var button = document.getElementById('addTagButton');
  var input = document.getElementById('customTagInput');
  if (isLoading) {
    button.classList.add('loading');
    button.disabled = true;
    input.disabled = true;
    button.textContent = 'Adding';
  } else {
    button.classList.remove('loading');
    button.textContent = 'Add Tag';
    input.disabled = false;
    toggleAddButton(); // This will set the appropriate enabled/disabled state
  }
}

// Function to add a custom tag
function addCustomTag() {
  return _addCustomTag.apply(this, arguments);
} // Function to update the selected tags display
function _addCustomTag() {
  _addCustomTag = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee5() {
    var input, tagName, normalizedNewTag, tagExists, response, responseData, tagObject;
    return _regeneratorRuntime().wrap(function _callee5$(_context5) {
      while (1) switch (_context5.prev = _context5.next) {
        case 0:
          input = document.getElementById('customTagInput');
          tagName = input.value.trim();
          if (tagName) {
            _context5.next = 4;
            break;
          }
          return _context5.abrupt("return");
        case 4:
          // Check if tag already exists
          normalizedNewTag = normalizeTagText(tagName);
          tagExists = availableTags.some(function (tag) {
            return normalizeTagText(tag.description) === normalizedNewTag;
          });
          if (!tagExists) {
            _context5.next = 9;
            break;
          }
          showTagError();
          return _context5.abrupt("return");
        case 9:
          setAddButtonLoading(true);
          _context5.prev = 10;
          _context5.next = 13;
          return makeAuthenticatedRequest('https://ml-inf-svc-dev.eventellect.com/corpus-collector/api/tag', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              description: tagName
            })
          });
        case 13:
          response = _context5.sent;
          _context5.next = 16;
          return response.json();
        case 16:
          responseData = _context5.sent;
          if (!(response.status === 400)) {
            _context5.next = 21;
            break;
          }
          showTagError(responseData.detail);
          setAddButtonLoading(false);
          return _context5.abrupt("return");
        case 21:
          if (response.ok) {
            _context5.next = 23;
            break;
          }
          throw new Error('Failed to create tag');
        case 23:
          tagObject = {
            id: responseData.id,
            description: responseData.description
          }; // Add the new tag to availableTags
          availableTags.push(tagObject);

          // Store the tag object in custom properties
          Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              var props = result.value;
              var selectedTags = props.get('selectedTags') || [];
              if (typeof selectedTags === 'string') {
                selectedTags = JSON.parse(selectedTags);
              }

              // Add the new tag to selected tags if not already present
              if (!selectedTags.some(function (tag) {
                return tag.id === tagObject.id;
              })) {
                selectedTags.push(tagObject);
                props.set('selectedTags', JSON.stringify(selectedTags));
                props.saveAsync(function () {
                  // After saving, update both displays with the new state
                  updateSelectedTagsDisplay(selectedTags);
                  displayFilteredTags(availableTags);
                  input.value = '';
                  setAddButtonLoading(false);
                });
              }
            }
          });
          _context5.next = 33;
          break;
        case 28:
          _context5.prev = 28;
          _context5.t0 = _context5["catch"](10);
          console.error('Error creating tag:', _context5.t0);
          setAddButtonLoading(false);
          showTagError('Failed to create tag. Please try again.');
        case 33:
        case "end":
          return _context5.stop();
      }
    }, _callee5, null, [[10, 28]]);
  }));
  return _addCustomTag.apply(this, arguments);
}
function updateSelectedTagsDisplay(selectedTags) {
  var selectedTagsContainer = document.getElementById('selectedTags');
  selectedTagsContainer.innerHTML = selectedTags.map(function (tag) {
    return "\n        <span class=\"selected-tag\" data-tag-id=\"".concat(tag.id, "\">\n            ").concat(tag.description, "\n            <span class=\"remove-tag\" onclick=\"removeTag('").concat(tag.id, "')\">&times;</span>\n        </span>\n    ");
  }).join('');
}

// Function to show loading state in tag list
function showTagListLoading() {
  var tagList = document.getElementById('tagList');
  tagList.innerHTML = '<div class="loading-spinner">Loading tags</div>';
}
window.commands = __webpack_exports__;
/******/ })()
;
//# sourceMappingURL=taskpane.js.map