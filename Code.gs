/**
* @OnlyCurrentDoc  Limits the script to only accessing the current document.
*/

//////////////////////////////////////////////////////////////////////////////////// globals

var SIDEBAR_TITLE = 'Свежий Взгляд';

var options = {};
//  wordcount_use_coefficient: 0, // 0..100
var max_wc_size = 100; // words... what a shame

var start_time;

//////////////////////////////////////////////////////////////////////////////////// Apps Script black magic

/**
* Adds a custom menu with items to show the sidebar and dialog.
*
* @param {Object} e The event parameter for a simple onOpen trigger.
*/
function onOpen(e) {
  DocumentApp.getUi()
  .createAddonMenu()
  .addItem('Проверить', 'showSidebar')
  .addToUi();
}

/**
* Runs when the add-on is installed; calls onOpen() to ensure menu creation and
* any other initializion work is done immediately.
*
* @param {Object} e The event parameter for a simple onInstall trigger.
*/
function onInstall(e) {
  onOpen(e);
}

/**
* Opens a sidebar. The sidebar structure is described in the Sidebar.html
* project file.
*/
function showSidebar() {
  var sb = HtmlService.createTemplateFromFile('Sidebar')
  .evaluate()
  .setTitle(SIDEBAR_TITLE)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(sb);
}

//////////////////////////////////////////////////////////////////////////////////// linguistic tables

var sim_ch_stride = 34;
var sim_ch = [                   /* letters' similarity map; order as in Unicode */
  /*а б в г д е ж з и й к л м н о п р с т у ф х ц ч ш щ ъ ы ь э ю я . ё */
    9,0,0,0,0,1,0,0,1,0,0,0,0,0,2,0,0,0,0,1,0,0,0,0,0,0,0,1,0,1,1,2,0,1 , /* а */
    0,9,1,0,0,0,0,0,0,0,0,0,0,0,0,3,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* б */
    0,1,9,1,0,0,0,0,0,0,0,1,1,1,0,1,0,0,0,1,3,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* в */
    0,0,1,9,0,0,3,0,0,0,3,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0 , /* г */
    0,0,0,0,9,0,0,1,0,0,0,0,0,0,0,0,0,1,3,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0 , /* д */
    1,0,0,0,0,9,0,0,2,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,2,1,1,0,9 , /* е */
    0,0,0,3,0,0,9,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,3,0,0,0,0,0,0,0,0 , /* ж */
    0,0,0,0,1,0,3,9,0,0,0,0,0,0,0,0,0,3,1,0,0,0,3,1,1,1,0,0,0,0,0,0,0,0 , /* з */
    1,0,0,0,0,2,0,0,9,3,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,2,0,1,1,1,0,2 , /* и */
    0,0,0,0,0,0,0,0,2,9,0,1,1,1,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* й */
    0,0,0,3,0,0,0,0,0,0,9,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0 , /* к */
    0,0,1,0,0,0,0,0,0,1,0,9,1,1,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* л */
    0,0,1,0,0,0,0,0,0,1,0,1,9,3,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* м */
    0,0,1,0,0,0,0,0,0,1,0,1,3,9,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* н */
    2,0,0,0,0,1,0,0,1,0,0,0,0,0,9,0,0,0,0,1,0,0,0,0,0,0,0,1,0,1,1,1,0,1 , /* о */
    0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,9,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* п */
    0,0,0,0,0,0,0,0,0,1,0,1,1,1,0,0,9,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0 , /* р */
    0,0,0,0,1,0,0,3,0,0,0,0,0,0,0,0,0,9,1,0,0,0,3,1,0,0,0,0,0,0,0,0,0,0 , /* с */
    0,0,0,0,3,0,0,1,0,0,0,0,0,0,0,0,0,1,9,0,0,0,1,1,0,0,0,0,0,0,0,0,0,0 , /* т */
    1,0,1,0,0,1,0,0,1,0,0,0,0,0,1,0,0,0,0,9,0,0,0,0,0,0,0,1,0,1,2,1,0,1 , /* у */
    0,1,3,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,9,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* ф */
    0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,0,0,0,0,9,0,1,0,0,0,0,0,0,0,0,0,0 , /* х */
    0,0,0,0,1,0,0,3,0,0,0,0,0,0,0,0,0,3,1,0,0,0,9,0,0,0,0,0,0,0,0,0,0,0 , /* ц */
    0,0,0,0,0,0,3,1,0,0,0,0,0,0,0,0,0,1,1,0,0,1,0,9,3,3,0,0,0,0,0,0,0,0 , /* ч */
    0,0,0,0,0,0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,9,3,0,0,0,0,0,0,0,0 , /* ш */
    0,0,0,0,0,0,3,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,9,0,0,0,0,0,0,0,0 , /* щ */
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,9,0,3,0,0,0,0,0 , /* ъ */
    1,0,0,0,0,1,0,0,2,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,9,0,1,1,1,0,1 , /* ы */
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,0,9,0,0,0,0,0 , /* ь */
    1,0,0,0,0,3,0,0,1,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,9,1,1,0,3 , /* э */
    1,0,0,0,0,1,0,0,1,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,0,0,0,1,0,1,9,1,0,1 , /* ю */
    2,0,0,0,0,1,0,0,1,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,1,1,9,0,1 , /* я */
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 , /* . */
    1,0,0,0,0,9,0,0,2,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,1,0,2,1,1,0,9   /* ё */
  /*а б в г д е ж з и й к л м н о п р с т у ф х ц ч ш щ ъ ы ь э ю я . ё */
];

var inf_letters_stride = 2;
var inf_letters  = [  /* quantity of information in letters */
  /* relative, average = 1000 */
  /*by itself - in the beginning of a word */
    802,    959 ,  /* а */
   1232,   1129 ,  /* б */
    944,    859 ,  /* в */
   1253,   1193 ,  /* г */
   1064,    951 ,  /* д */
    759,   1232 ,  /* е */
   1432,   1432 ,  /* ж */
   1193,    993 ,  /* з */
    802,    767 ,  /* и */
   1329,   1993 ,  /* й */
   1032,    929 ,  /* к */
    967,   1276 ,  /* л */
   1053,    944 ,  /* м */
    848,    711 ,  /* н */
    695,    853 ,  /* о */
   1088,    454 ,  /* п */
    929,   1115 ,  /* р */
    895,    793 ,  /* с */
    848,   1002 ,  /* т */
   1115,   1129 ,  /* у */
   1793,   1022 ,  /* ф */
   1259,   1329 ,  /* х */  /* [0] manually decreased! was 1359 */
   1593,   1393 ,  /* ц */
   1276,   1212 ,  /* ч */
   1476,   1012 ,  /* ш */
   1676,   1676 ,  /* щ */
   1993,   3986 ,  /* ъ */
   1193,   3986 ,  /* ы */
   1253,   3986 ,  /* ь */
   1676,   1232 ,  /* э */
   1476,   1793 ,  /* ю */
   1159,    967 ,  /* я */
  0, 0,
   1300,   1900    /* ё */  /*set manually - dk*/
];


var exceptions_voc = {       /* exceptions vocabulary */
    "белым бело": true,
    "больше больше": true,
    "больше более": true,
    "более больше": true,
    "больше меньше": true,
    "бы бы": true,
    "бы был": true,
    "бы была": true,
    "бы были": true,
    "бы было": true,
    "бы вы": true,
    "был бы": true,
    "была бы": true,
    "были бы": true,
    "было бы": true,
    "вины виноватый": true,
    "волей неволей": true,
    "время времени": true,
    "всего навсего": true,
    "вы бы": true,
    "даже уже": true,
    "друг друга": true,
    "друг друге": true,
    "друг другом": true,
    "друг другу": true,
    "дурак дураком": true,
    "если если": true,
    "звонка звонка": true,
    "или или": true,
    "как так": true,
    "конце концов": true,
    "корки корки": true,
    "кто что": true,
    "либо либо": true,
    "мало помалу": true,
    "меньше больше": true,
    "начать сначала": true,
    "не на": true,
    "не не": true,
    "не ни": true,
    "него нет": true,
    "ни на": true,
    "ни не": true,
    "ни ни": true,
    "но на": true,
    "но не": true,
    "но ни": true,
    "новые новые": true,
    "объять необъятное": true,
    "одному тому": true,
    "полным полно": true,
    "постольку поскольку": true,
    "так как": true,
    "тем чем": true,
    "то то": true,
    "тогда когда": true,
    "ха ха": true,
    "чем тем": true,
    "тем чем": true,
    "что то": true,
    "чуть чуть": true,
    "шаг шагом": true,
    "этой что": true,
    "этот что": true,
  };

// separate dict with first words only, for quick check
var exceptions_voc_first = {};
for (var key in exceptions_voc) {
   if (exceptions_voc.hasOwnProperty(key)) { 
      exceptions_voc_first[key.split(" ")[0]] = true;
   }
}  

function checkvoc (w1, w2) {  
  if (exceptions_voc_first[w1]) {
      return exceptions_voc[w1+" "+w2];
  }
  return false;
}

function weigh_sep(prec_sep) {
  var res = prec_sep.new_element? 8 : 0;
  var seen_spaces = false;
  var new_sentence = prec_sep.new_element; // new sentence starts if new element, but also after punctuation below
  for (var i = 0, len = prec_sep.sep_str.length; i < len; i++) {
    switch (prec_sep.sep_str[i]) {
      case ' ':
        if ( !seen_spaces ) {
          res ++;
          seen_spaces = true;
        }
        break;
        
      case ',':
      case '/':
        res += 2;
        break;
        
      case '.':
      case '?':
      case '!':
        res += 4;
        new_sentence = true;
        break;
        
      case ')':
      case '(':
      case '[':
      case ']':
      case '<':
      case '>':
        res += 3;
        break;
        
      case '"':
      case '\'':
      case '«':
      case '»':
      case '“':
      case '”':
      case '„':
      case '“':
        res += 3;
        new_sentence = true;
        break;
        
      case '—':
      case '—':
        res += 3;
        break;        
        
      case '-':
        if ( seen_spaces )
          res += 3;
        else
          res ++;
        break;
        
      default:
        // usually a one-letter word
        res += 3;
        break;
    }
  }
  return {
    weight: res,
    new_sentence: new_sentence
  }
}

//////////////////////////////////////////////////////////////////////////////////// core comparer of words

  // internal subfunctions

  /*
   * Psych. importance of the word x ch. long big for small words, then slowly
   * lagging  behind the real length
   */
  function implen (x)
  {
    if (x == 2)
      return 5.0; // not 5 to avoid polymorphism in return value
    return (x - ((x - 1) * (x - 1) / 36) + (4.1 / x));
  }
  
  /*
  * Calculates: average quantity of information in the letters common to
  * the two words; total quantity of information in differing letters of the
  * two words.
  */
  function infor_same_diff (a, b)
  {
    var count_same = 0;
    var res_same = 0;
    var count_diff = 0;
    var res_diff = 0;
    
    for (var i = 0, n = a.length; i < n; i++) {
      var bp = b.indexOf(a[i]);
      if (bp != -1) { /* bipresent letters */
        if (i == 0 && bp == 0)
          res_same += inf_letters [a[i] * inf_letters_stride + 1 ];  /* beginning of the word */
        else
          res_same += inf_letters [a[i] * inf_letters_stride ];  /* elsewhere */
        count_same ++;
      }
      else /* letters in a only */
      {
        if (i == 0)
          res_diff += inf_letters [a[i] * inf_letters_stride + 1 ]; /* in the beginning of the word */
        else
          res_diff += inf_letters [a[i] * inf_letters_stride ]; /* elsewhere */
        count_diff ++;        
      }
    }

    for (var i = 0, n = b.length; i < n; i++) {              /* letters in b only */
      var ba = a.indexOf(b[i]);
      if (ba == -1) {
        if (i == 0)
          res_diff += inf_letters [b[i] * inf_letters_stride + 1 ]; /* in the beginning of the word */
        else
          res_diff += inf_letters [b[i] * inf_letters_stride ]; /* elsewhere */
        count_diff ++;
      }
    }
    
    return [count_same? (res_same / count_same) : 0.0, count_diff ? res_diff : 0];
  }


  // simwords proper
function simwords (a, b) //// core of the program: a score of how similar two words look and sound
{  
  var dissimilarity_threshold = 24000; /* how much total information in 
					   differing letters reduces res
					   to zero; larger values make res
					   more tolerant to differences */
  
  var infor_same_diff_value = infor_same_diff (a, b);
  
  if (infor_same_diff_value[1] >= dissimilarity_threshold) 
 			/* too many too rare dissimilar letters? */
    return (0);
  
  // speed up
  var alen = a.length;
  var blen = b.length;

  var rever = false;
  if ( alen > blen ) {  /* swap strings so a is always the shortest */
    var tmp = a;
    a = b;
    b = tmp;
    tmp = alen;
    alen = blen;
    blen = tmp;
    rever = true;
  }
  
  var reciproc_3_alen = 1.0 / (3 * alen);
  var reciproc_3_blen = 1.0 / (3 * blen);

  var res = 0; /* value to be returned */
  var resa = 0, partlen = 1;
  var dist;
  // speed up by using local vars
  
  for ( partlen = 1; partlen <= alen; partlen ++, resa = 0 ) {

    for ( var ta = 0; //a -> logical;  
        (alen - ta) >= partlen;  
        ta ++) {

      for ( var tb = 0; // b -> logical;  
          (blen - tb) >= partlen;  
          tb ++) {
            
        var tx = ta, ty = tb, prir = 0;
        for ( ; // prir = 0, tx = parta, ty = tb 
            tx < ta + partlen; 
            tx ++, ty ++)
          prir += sim_ch [ a[tx] * sim_ch_stride + b[ty] ];
          if ( !prir ) 
            continue;
          if ( ta > 0 )
            prir -= (prir * ta) * reciproc_3_alen;
          if ( tb > 0 )
            prir -= (prir * tb) * reciproc_3_blen;
        
          dist = rever ? 
              (blen - (tb + partlen)) + ta
               : 
              (alen - (ta + partlen)) + tb;
        
          if ( dist < 3 )
            prir += ((prir * (2 - dist)) * 0.333);

          if (prir > resa)
            resa = prir;         
      }
    }
    if (resa > partlen * 6) {
      prir = resa;
      dist = (alen + blen) * 0.375 + 1; // * 3 / 8
      res += resa + prir * (partlen - (dist < alen? dist : alen)) / (2 * dist);
    }
  }
  
  for (partlen = 1, resa = 0;  partlen <= alen;  partlen ++)
    resa += 9 * partlen;

  res = ((res * infor_same_diff_value[0]) / resa);
	      /* allowing for the info contained in the common letters */
  res = (res * (dissimilarity_threshold - infor_same_diff_value[1]))/dissimilarity_threshold;
	      /* decreasing by a coefficient depending on infor_diff */

  res -= (res * (blen - alen)) / (2 * blen);
	      /* decreasing if words are too different in length */

  return (res * alen * blen / (implen (alen) * implen (blen)));
	      /* finally, taking into account the psychological length */
}
  
/////////////////////////////////////////////////////////////////////////////////// worder

function make_worder(source) { //// an object that feeds us words and separators; one of the two places that touches the Google Doc API (read only), the other being context::paint  
  var elements;
  if ('getRangeElements' in source) { // source can be either selection or Body; this is a selection
    elements = source.getRangeElements(); // all elements that selection touches; even if a paragraph is partially selected, we check it all
    for (var i = 0; i < elements.length; i++) {
      elements[i] = elements[i].getElement().asText(); // make an array of Text elements that have findText
    }
  }
  else 
  { // single Body element (all document); it already has findText
    elements = [source];
  }
    
  return {    
    pattern: "[А-ЯЁа-яё][а-яё]+", // cannot search for single letters because of https://code.google.com/p/google-apps-script-issues/issues/detail?id=2770, so at least two russian letters
    elements: elements,
    last_found_word: null,
    last_i: 0,
    new_element: true,
    
    _cache: new Array(5000),
    _cache_length: 0,
    _index: 0,
    
    reset: function() { this._index = 0; }, // so it can be reused, will return words from cache

    get_word: function() { //// returns the next word, its indices in the source so it can be highlighted, and the sep before it
      
      if (this._cache_length > this._index) {
        return this._cache[this._index++];
      }
      
      this.last_found_word = this.elements[0].findText(this.pattern, this.last_found_word);
      while (this.last_found_word == null) { // reached end of an element, go to next one
        this.elements.shift(); // pop the first in the list
        this.new_element = true;
        if (this.elements.length == 0) { // end of elements
          return null;
        }
        this.last_found_word = this.elements[0].findText(this.pattern, this.last_found_word);
      }
      
      var last_found_word = this.last_found_word;
      var new_element = this.new_element;
      var last_i = this.last_i
      var word_element = last_found_word.getElement().asText();
      var text = word_element.getText();
      var start = last_found_word.getStartOffset();
      var end = last_found_word.getEndOffsetInclusive();
      var r = {
        word_element: word_element,
        word_element_text: text,
        word_str: text.slice(start, end + 1),
        start: start,
        end: end,
        prec_sep: {
          new_element: new_element,
          sep_str: text.slice(last_i, start)
        }
      };
      this.new_element = false;
      this.last_i = end + 1;

      if (!(this._cache_length <= this._index)) {
        this._cache[this._index] = r;
        this._cache_length ++;
      }
      this._index ++;

      return r;
    }
  }
}

/////////////////////////////////////////////////////////////////////////////////// wordcount - TODO, now disabled

/*
TODO for wordcount:
1. add controls for it: compile button, highlight button (use another color!), usage percentage scroller (or just checkbox?); latter two are disabled until you compile it once
2. compile it always from entire document, by grabbing all its text and parsing it in the script without those findText calls
3. store wordcount persistently (cache service?)
4. consider NOT using stemmer, just plain wordform count (depends on speed)
5. combine with a real frequency word list from a big modern corpus
*/

function get_stem_len(l) {
  switch(l) {
    case 2: return 2;
    case 3: return 3;
    case 4: return 3;
    case 5: return 3;
    case 6: return 4;
    case 7: return 5;
    default: return l - 3;
  }
}

function wc_same(s1, s2) { //// extremely primitive stemmer that just disregards endings when comparing words
  s1 = s1.toLowerCase();
  //s2 = s2.toLowerCase(); // skip, because second string always comes from dict keys, always lowercase
  var stem_len = get_stem_len(Math.max(s1.length, s2.length));
  return s1.slice(0, stem_len) == s2.slice(0, stem_len);
}


function make_wordcount(worder) {
  wc = {
    dict: {},
    get: function(w) { //// get count for a word
      if (w in this.dict) { // optimization: if literally this word is in dict
        return this.dict[w];
      }
      for (var key in this.dict) { // otherwise search by our comparer
        if (wc_same(w, key)) {
          return this.dict[key];
        }
      }
      return null;
    },
    increment: function(w) { //// increment count for a word
      if (w in this.dict) { // optimization: if literally this word is in dict
        this.dict[w] ++;
        return;
      }        
      for (var key in this.dict) { // otherwise search by our comparer
        if (wc_same(w, key)) {
          this.dict[key] ++;
          return;
        }
      }
      this.dict[w.toLowerCase()] = 1;
    },
    max: function() { //// return max count 
      var m = 0;
      for (var key in this.dict) {
        if (this.dict[key] > m)
          m = this.dict[key];
      }      
      return m;
    }
  };
  var worder = worder;
  var count = 0;
  while ((word = worder.get_word()) != null && count < max_wc_size) {
    wc.increment(word.word_str);
    count ++;
  }
  
  // convert counts to coefficients: from 1000 for 1-occurrence words down for more common words
  for (var key in wc.dict) {
	if (wc.dict[key] == 1)
		wc.dict[key] = 1000;
	else {
		var tmp = wc.dict[key] / Object.keys(wc.dict).length;
	    /*
		* Decreasing the second 8.0 will sharpen the dependance
		* on wordcount
		*/
		wc.dict[key] = 1 - Math.log(tmp) * 1000 / ((8.0 + options.wordcount_use_coefficient / 8.0) * Math.log(2));
		if (wc.dict[key] > 1000)
			wc.dict[key] = 1000;
	}
  }
  
  //DocumentApp.getUi().alert(JSON.stringify(wc.dict));
  return wc;
}

/////////////////////////////////////////////////////////////////////////////////// context

function make_context (wc, worder) {
  return {
    worder: worder,
    twosigmasqr_reciprocal: 1 / (2 * Math.pow(options.context_size * 4, 2)),
    wc: wc,
    _queue: [], 
    _bads: [],
    total_badness: 0,
    
    shift: function() { //// reads a new word, adds it to queue, trims queue if too long
      function to_lnum(w) { //// converts word to array of indices into linguistic tables, for speed
        var n = w.length;
        var r = new Array(n);
        for (var i = 0; i < n; i++) {
          r[i] = (w[i]).charCodeAt(0) - 1072; // 1072 = 'а'.charCodeAt(0););
        };
        return r;
      };
      var word = this.worder.get_word();
      if (word) {
        var sep = weigh_sep(word.prec_sep);
        var word_str = word.word_str.toLowerCase();
        this._queue.push({
          is_proper: /^[А-Я][а-я]/.test(word.word_str) && !sep.new_sentence, // is Capitalized but this is not start of new sentence
          word: word_str,
          word_lnum: to_lnum(word_str),
          word_element: word.word_element,
          word_element_text: word.word_element_text, // so that we can identify elements - there's no way to get element id or hash from google
          start: word.start,
          end: word.end,
          sep: sep.weight,
          new_sentence: sep.new_sentence
        });
        if (this._queue.length > options.context_size) { // trim from start
          this._queue.shift();
        }
        return true;
      }
      return false; // end of input
    },
    
    check: function() { //// checks the head of the queue against the other words in queue, remembers any bads found
      var exp = Math.exp;
      var current = this._queue[this._queue.length - 1];
      if (options.exclude_proper_names && current.is_proper)
        return;
      for (var i = 0, n = this._queue.length - 1; i < n; i++) {
        var other = this._queue[i];
        if (options.exclude_proper_names && other.is_proper)
          continue;
        if (checkvoc (other.word, current.word)) /* an exception? */
          continue;
        var sim = simwords ( other.word_lnum, current.word_lnum ); // compare ith word with the last (current) one
        //DocumentApp.getUi().alert(other.word+"|"+current.word+"|"+sim);  

        var dist = 0; // calculate psychological distance between the words
        for (var j = i + 1, nn = this._queue.length; j < nn; j ++) {
          dist += this._queue[j].sep; // add all separators
          if (j < nn - 1)
             dist += this._queue[j].word.length * 0.333 + 1; // words themselves are also separators, count in their length
        }
        if (options.wordcount_use_coefficient) { // increase dist for frequent words
			dist *= 2000; // because we divide by sum of coefficients, each 1000 for 1-occurrence words
			dist /= wc.get(this._queue[i].word) + wc.get(this._queue[this._queue.length - 1].word);
        }
        
        var dal = exp ((- dist * dist) * this.twosigmasqr_reciprocal);
        var badness = sim * dal;
	
        if ( badness > options.sensitivity_threshold ) {
          // add to list of paintables
          this._bads.push({
            badness: badness,
            sim: sim,
            dist: dist,
            other: other,
            current: current
          });
          this.total_badness += badness;
        }
      }
	},
      
    run: function() { //// keeps shifting and checking until there are words in worder, returns stats
      var words_checked = 0;      
      var broken = false;
      while (this.shift()) {
        this.check();
        words_checked ++;
        // do we have time yet? google limits scripts to 5 minutes
        var now = new Date();
        if (now.getTime() - start_time.getTime() > 240000) { // 60000 * 4 minutes
          broken = true;
          break;         
        }
      }
      var average_badness = (words_checked? this.total_badness/words_checked : 0);
      average_badness = Math.round(average_badness * 100) / 100;
      return "Готово.<br/>"+
             "Слов: "+words_checked+(broken?" (сколько успел, извините)":"")+"<br/>"+
             "Плохих пар: "+this._bads.length+"<br/>"+
             "Средняя плохость на слово: "+average_badness;
    },
      
    paint: function() { //// paints all bads in the document with colors corresponding to badness; interfaces to google doc to do so
      var already_painted = {};
      for (var i = 0; i < this._bads.length; i++) {
        var bad = this._bads[i];
        var badness = bad.badness;
        if (badness > 2000) badness = 2000; 
        
        var keyother = ""+bad.other.start+";;"+bad.other.end+";;"+bad.other.word_element_text;
        var keycurrent = ""+bad.current.start+";;"+bad.current.end+";;"+bad.current.word_element_text;
        
        var level = (badness - options.sensitivity_threshold)/(2000 - options.sensitivity_threshold);
        
        if (keyother in already_painted) {
          highlight_color_counter = already_painted[keyother][1];
        }
        if (keycurrent in already_painted) {
          highlight_color_counter = already_painted[keycurrent][1];
        }
        
        var color = get_color_for_tint(level, highlight_color_counter);

        // paint this one only if it was not yet painted by another pair, or if it was that badness was less than the current
        if (!(keyother in already_painted) || already_painted[keyother][0] < badness) {
          bad.other.word_element.setBackgroundColor(bad.other.start, bad.other.end, color);
          already_painted[keyother] = [badness, highlight_color_counter];
        }
        if (!(keycurrent in already_painted) || already_painted[keycurrent][0] < badness) {
          bad.current.word_element.setBackgroundColor(bad.current.start, bad.current.end, color);        
          already_painted[keycurrent] = [badness, highlight_color_counter];
        }
        
        highlight_color_counter++;
      }
    }
  }  
}

/////////////////////////////////////////////////////////////////////////////////// color

function decimalToHex(d, padding) {
  var hex = Number(d).toString(16);
  while (hex.length < padding) {
      hex = "0" + hex;
  }
  return hex;
}

function get_rgb_color(r, g, b)
{
  return "#" + decimalToHex(r,2) + decimalToHex(g,2) + decimalToHex(b,2);
}

var no_color        = {r: 255, g: 255, b: 255};

function get_color_for_tint(level, color_counter) { // level: 0..1, color_counter: which of the three tints
  var highlight_colors = [{r: 235, g: 20, b: 255},  // purplish
                          {r: 255, g: 130, b: 25},  // reddish/yellowish
                          {r: 165, g: 70, b: 255},  // bluish
                         ];
  var highlight_color = highlight_colors[color_counter % 3];
  
  level += (1 - level) * 0.1 // hike lower values to avoid too pale highlights

  var r = Math.round(no_color.r + (highlight_color.r - no_color.r) * level);
  var g = Math.round(no_color.g + (highlight_color.g - no_color.g) * level);
  var b = Math.round(no_color.b + (highlight_color.b - no_color.b) * level);
  return get_rgb_color(r, g, b);
}

var highlight_color_counter = 0;

function get_color(level) { //// get highlight color for (float) level of badness
  highlight_color_counter++;
  return get_color_for_tint(level, highlight_color_counter);
}

/////////////////////////////////////////////////////////////////////////////////// main

function fresheye(source, sensitivity_threshold, context_size, exclude_proper_names) {
  if (!source)
    return "Ничего не выделено";
  
  start_time = new Date(); // must watch time and stop before it runs out of time limit for Apps Script - 6 minutes :(
  
  options.sensitivity_threshold = sensitivity_threshold;
  options.context_size = context_size;
  options.exclude_proper_names = exclude_proper_names;

  var worder = make_worder(source);
  
  var wc;
  if (options.wordcount_use_coefficient > 0) { // FIXME: now never used
    wc = make_wordcount(worder);
    worder.reset();
  }
  
  var ctx = make_context(wc, worder); 
  var diag = ctx.run();
  ctx.paint();
  return diag;
}

//////////////////////////////////////////////////////////////////////////////// external API

function fresheye_document(sensitivity_threshold, context_size, exclude_proper_names) {

  var doc = DocumentApp.getActiveDocument();
  var source = doc.getBody();
    
  return fresheye(source, sensitivity_threshold, context_size, exclude_proper_names);
}

function fresheye_selection(sensitivity_threshold, context_size, exclude_proper_names) {
  
  var doc = DocumentApp.getActiveDocument();
  var source = doc.getSelection();
  
  return fresheye(source, sensitivity_threshold, context_size, exclude_proper_names);
}

function clear_document_or_selection() {
  var doc = DocumentApp.getActiveDocument();
  if (doc.getSelection()) {
    var sel = doc.getSelection();
    elements = sel.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var text = elements[i].getElement().asText();
      if (text.getText().length > 0)
        text.setBackgroundColor(0, text.getText().length - 1, get_rgb_color(no_color.r, no_color.g, no_color.b));
    }
  } else {
    var text = doc.getBody().editAsText();
    text.setBackgroundColor(0, text.getText().length - 1, get_rgb_color(no_color.r, no_color.g, no_color.b));
  }
  
  return true; 
}
