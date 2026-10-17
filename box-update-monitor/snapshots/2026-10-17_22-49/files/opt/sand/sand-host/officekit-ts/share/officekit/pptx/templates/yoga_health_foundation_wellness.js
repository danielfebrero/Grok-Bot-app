#!/usr/bin/env node
/**
 * "Yogaes - Healthy & Yoga Foundation" - 45 slide deck, 13.333 x 7.5 in.
 * Rebuilt from scratch with pptxgenjs: every shape, colour and string below
 * is a plain literal. Photographs of the original are replaced by labelled
 * placeholder rectangles (see `photo()`).
 */
'use strict';
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette
   Theme colours of the original deck (accent1..6 with luminance variants).  */
const BG = '424D46';      // page background   - accent6 @ 50% luminance
const CREAM = 'FEF4DF';   // accent1           - pale body text / petals
const CREAM2 = 'FDE4B1';  // accent1 @ 90%     - monstera leaves
const SAND = 'FDD5A1';    // accent2           - big background blobs
const SAND2 = 'FEE6C7';   // accent2 @ 40% lum - card fills
const SAND3 = 'FCC378';
const SAND4 = 'F2BF86';
const ORNG = 'EA9536';    // accent3           - headings, buttons, ferns
const ORNG2 = 'FBA83C';   // accent3 @ 75% lum
const GOLD = 'E79F07';    // accent1 @ 50% lum - "YOGAES" wordmark
const SAGE = '86978C';    // accent6
const SAGE2 = '637369';   // accent6 @ 75% lum - body copy on cream
const BROWN = '653A02';
const INK = '3B3838';
const W = 'FFFFFF';
const FONT = 'Baloo 2';

/* ------------------------------------------------------- repeated copy
   The deck reuses the same handful of filler sentences everywhere.        */
const L = {
  allAspectsOf: 'All aspects of physicality in the universe are cyclical. Planets are going around the sun, the solar system is moving',
  allAspectsOf2: 'All aspects of physicality in the universe are cyclical. Planets are going around the sun, the solar system is moving, and everything in the galaxy and the cosmos is cyclical. ',
  allAspectsOf3: 'All aspects of physicality in the universe are cyclical. Planets are going around the sun, the solar system is moving, and everything in the galaxy and the cosmos is cyclical. The more you are identified with your physical system, the more cyclical you are. Your experiences and the process of life is cyclical. If you watch carefully enough, even the situations that you face in your life come in cycles.',
  balanceBody: 'Balance Body',
  beautifulBodyAnd: 'Beautiful Body And ',
  benefitOfYoga: 'Benefit Of Yoga',
  chartsPlans: 'Charts Plans',
  chooseYourFavori: 'Choose Your Favorites Pricing Plans',
  expertService: 'Expert Service',
  expertStaff: 'Expert Staff',
  getTheDiscount: 'Get The 20% Discount',
  healthYogaFounda: ' Health & Yoga Foundations',
  healthyLife: 'Healthy Life',
  healthyYogaFound: 'Healthy & Yoga Foundation',
  meditationTime: 'Meditation Time',
  mentalHealth: 'Mental Health',
  mindBuiltOn: ' Mind Built On You',
  peacefulMind: 'Peaceful Mind',
  peacefulMinds: 'Peaceful Minds',
  professionalYoga: 'Professional Yoga Trainer',
  published: 'PUBLISHED : ',
  qualifiedService: 'Qualified Services For Teach You About Yoga',
  s: 'HEALTH_YOGA_FOUNDATION_',
  startFocusOn: 'Start Focus On Routine Yoga Everyday  2 Hours',
  theFeatures: 'The Features',
  theMoreYou: 'The more you are identified with your physical system, the more cyclical you are. Your experiences and the process of life is cyclical. If you watch carefully enough, even the situations that you face in your life come in cycles.',
  todayModernPhysi: 'Today, Modern Physics Is Proving To You That As You Sit Here, Every Subatomic Particle In Your Body Is In Constant Transaction With Everything Else In The Existence. If This Transaction Stops, You Will Cease To Exist. So, Yoga Means To Know The Union Of Existence By Experience.',
  todayModernPhysi2: 'Today, Modern Physics Is Proving To You That As You Sit Here, Every Subatomic Particle In Your Body Is In Constant Transaction With Everything Else In The Existence.',
  todayModernPhysi3: 'Today, Modern Physics Is Proving To You That As You Sit Here, Every Subatomic Particle In Your Body Is In Constant Transaction With Everything Else In The Existence. ',
  todayModernPhysi4: 'Today, Modern Physics Is Proving To You That As You Sit Here, Every Subatomic Particle In Your Body Is In Constant Transaction With Everything Else In The Existence',
  weHaveSpecial: 'We Have Special Events For All Member Of Yoga Meditations',
  weHaveSpecial2: 'We Have Special Mobile Apps For All Member Yoga',
  yogaMeditationPr: 'Yoga Meditation Program For Healthy ',
  yogaProgram: 'Yoga Program',
  youWillCease: 'You Will Cease To Exist. So, Yoga Means To Know The Union Of Existence By Experience.',
};

/* --------------------------------------------------------------- artwork
   Decorative leaves, ferns, sparkles and organic blobs of the template,
   traced as normalised outlines (0..1 of the shape box) and drawn with
   pptxgenjs `custGeom`.  'M x,y' = move, 'L x,y' = line, 'Z' = close.     */
const ART = {
  blobBig: 'M.568,0 L.451,.015 L.369,.064 L.324,.113 L.174,.324 L.067,.448 L.024,.527 L.003,.609 L.005,.732 L.047,.88 L.108,1 L1,1 L.949,.92 L.923,.812 L.929,.729 L.978,.588 L.996,.491 L.992,.39 L.947,.243 L.858,.118 L.776,.056 L.677,.015 Z',
  blob1: 'M0,0 L.975,0 L.995,.055 L.999,.113 L.969,.181 L.907,.225 L.835,.245 L.684,.245 L.625,.273 L.604,.306 L.589,.366 L.597,.582 L.582,.628 L.561,.658 L.518,.691 L.466,.713 L.225,.745 L.145,.775 L.091,.83 L0,1 Z',
  blobCorner: 'M0,0 L1,0 L.945,.106 L.907,.146 L.859,.178 L.8,.199 L.618,.226 L.532,.251 L.46,.288 L.4,.338 L.353,.403 L.316,.488 L.256,.745 L.211,.848 L.175,.898 L.116,.952 L.058,.983 L0,1 Z',
  blob2: 'M.968,0 L.032,0 L.014,.008 L0,.046 L0,.954 L.014,.992 L.032,1 L.968,1 L.986,.992 L1,.954 L1,.046 L.991,.013 Z',
  blob3: 'M.028,1 L.008,.988 L0,.959 L0,.041 L.013,.007 L.028,0 L.987,.007 L1,.041 L1,.959 L.987,.993 Z',
  blobSoft: 'M1,0 L1,1 L.044,1 L.012,.949 L0,.9 L.009,.855 L.039,.816 L.139,.767 L.333,.746 L.386,.707 L.404,.628 L.374,.457 L.405,.379 L.498,.323 L.747,.275 L.833,.238 L.886,.178 L.95,.042 Z',
  blobSide: 'M.998,.318 L.989,.464 L.96,.562 L.913,.656 L.778,.823 L.692,.89 L.597,.944 L.445,.993 L.349,1 L.26,.988 L.181,.958 L.113,.911 L.038,.812 L.009,.729 L0,.634 L.011,.536 L.061,.391 L.148,.257 L.222,.178 L.308,.11 L.403,.056 L.505,.018 L.604,.001 L.697,.004 L.781,.025 L.887,.089 L.941,.151 L.978,.228 Z',
  blobWave: 'M0,0 L1,0 L.947,.252 L.941,.646 L.907,.763 L.847,.854 L.795,.89 L.745,.899 L.561,.86 L.501,.863 L.437,.887 L.281,.979 L.215,.998 L.135,.99 L.064,.94 L0,.838 Z',
  blob4: 'M.944,.996 L1,.939 L.94,.955 L.497,.9 L.434,.877 L.389,.841 L.04,.037 L.044,0 L.003,.052 L.003,.084 L.355,.885 L.4,.921 L.457,.943 Z',
  blob5: 'M.492,.95 L.408,.92 L.353,.862 L.005,.064 L.002,.029 L.024,.007 L.078,.001 L.524,.059 L.606,.095 L.651,.146 L1,.964 L.976,.993 L.934,1 Z',
  blob6: 'M.49,.95 L.408,.92 L.356,.864 L.005,.062 L.002,.028 L.024,.006 L.076,.001 L.525,.059 L.605,.094 L.648,.143 L1,.965 L.977,.993 L.936,1 Z',
  blob7: 'M.489,.949 L.408,.92 L.356,.865 L.005,.061 L.002,.028 L.023,.006 L.075,.001 L.526,.059 L.604,.094 L.647,.142 L1,.965 L.977,.993 L.937,1 Z',
  monstera: 'M.629,.294 L.612,.309 L.568,.29 L.402,.138 L.273,.327 L.41,.415 L.447,.459 L.458,.497 L.384,.482 L.237,.405 L.108,.648 L.196,.672 L.231,.71 L.192,.725 L.09,.716 L0,.914 L.161,.948 L.253,.832 L.3,.8 L.321,.813 L.316,.851 L.198,.982 L.307,.997 L.388,.99 L.423,.954 L.5,.725 L.533,.692 L.548,.704 L.55,.738 L.501,.989 L.527,1 L.579,.99 L.712,.913 L.752,.701 L.784,.647 L.804,.674 L.81,.743 L.77,.92 L.883,.843 L.939,.776 L.983,.684 L.999,.609 L.986,.519 L.924,.439 L.818,.387 L.824,.274 L.806,.167 L.77,.083 L.708,.017 L.659,.001 L.581,.008 L.437,.081 L.597,.21 L.629,.262 Z M.652,.673 L.671,.707 L.664,.764 L.641,.813 L.608,.825 L.593,.792 L.6,.737 L.621,.687 Z M.476,.66 L.447,.748 L.414,.782 L.392,.78 L.386,.757 L.406,.716 Z M.602,.437 L.562,.436 L.51,.378 L.552,.372 L.591,.387 L.612,.412 Z',
  blobSide2: 'M.998,.318 L.989,.464 L.913,.656 L.778,.823 L.597,.944 L.445,.993 L.349,1 L.26,.988 L.18,.958 L.113,.911 L.038,.812 L.009,.729 L0,.634 L.011,.536 L.061,.391 L.148,.257 L.222,.178 L.403,.056 L.505,.018 L.604,.001 L.697,.004 L.781,.025 L.887,.089 L.941,.151 L.978,.228 Z',
  blob8: 'M.02,.986 L.53,.989 L.588,.975 L.637,.943 L.657,.913 L.902,.074 L.888,.027 L.823,0 L.939,.015 L.985,.039 L.999,.086 L.754,.926 L.734,.954 L.681,.987 L.621,1 Z',
  blob9: 'M.101,.997 L.029,.979 L.001,.925 L.267,.048 L.322,.013 L.382,0 L.911,.003 L.978,.027 L.998,.083 L.733,.952 L.678,.987 L.618,1 Z',
  blob10: 'M.098,.997 L.028,.979 L0,.927 L.269,.047 L.322,.013 L.38,0 L.914,.003 L.979,.026 L.998,.081 L.731,.953 L.678,.987 L.62,1 Z',
  blob11: 'M.096,.997 L.027,.979 L.001,.928 L.27,.046 L.323,.013 L.38,0 L.915,.003 L.98,.026 L.998,.079 L.73,.954 L.678,.987 L.621,1 Z',
  fernB: 'M.57,.356 L.433,.42 L.549,.327 L.636,.288 L.732,.193 L.928,.1 L1,0 L.666,.158 L.555,.311 L.387,.445 L.506,.236 L.525,.049 L.456,.146 L.386,.312 L.38,.451 L.28,.569 L.301,.395 L.289,.25 L.23,.116 L.217,.347 L.277,.574 L.232,.75 L.205,.679 L.187,.535 L.115,.403 L.077,.371 L.151,.666 L.232,.753 L.252,.901 L.116,.68 L0,.583 L.112,.751 L.244,.889 L.258,.999 L.271,.955 L.256,.826 L.421,.77 L.63,.669 L.685,.616 L.703,.55 L.64,.598 L.339,.712 L.248,.826 L.254,.682 L.422,.599 L.502,.537 L.682,.487 L.818,.418 L.616,.447 L.413,.518 L.345,.562 L.258,.662 L.29,.575 L.348,.498 L.424,.427 L.796,.358 L.976,.248 Z',
  blob12: 'M.334,0 L.634,.047 L.851,.163 L.979,.322 L1,.425 L.981,.525 L.899,.673 L.655,1 L0,1 L0,.057 L.165,.014 Z',
  leaf28: 'M.182,0 L1,0 L1,.993 L.804,.99 L.641,.925 L.558,.838 L.484,.638 L.429,.554 L.347,.495 L.11,.386 L.036,.329 L.002,.267 L.015,.18 L.081,.083 Z',
  leaf17: 'M.277,.948 L.436,.998 L.793,.986 L.958,.912 L.999,.815 L.96,.75 L.641,.554 L.513,.426 L.492,.389 L.504,.347 L.578,.222 L.184,0 L.036,.227 L0,.429 L.022,.548 L.067,.652 L.242,.854 L.242,.91 Z',
  leaf34: 'M.362,0 L.37,.029 L.391,.024 L.39,.051 L.412,.019 L.44,.048 L.427,.191 L.442,.126 L.467,.143 L.491,.115 L.514,.131 L.532,.093 L.557,.1 L.597,.174 L.591,.126 L.611,.141 L.639,.118 L.628,.065 L.665,.064 L.666,.095 L.684,.057 L.713,.098 L.753,.067 L.796,.123 L.823,.071 L.85,.099 L.834,.047 L.843,.058 L.867,.032 L.942,.048 L.965,.084 L.959,.05 L.986,.05 L.962,.149 L.947,.131 L.96,.207 L.938,.202 L.985,.263 L.956,.388 L.941,.372 L.912,.408 L.941,.518 L.931,.647 L.884,.52 L.903,.377 L.89,.318 L.881,.397 L.871,.361 L.849,.379 L.856,.454 L.835,.476 L.816,.454 L.771,.498 L.751,.623 L.772,.649 L.779,.629 L.803,.647 L.82,.756 L.791,.903 L.771,.902 L.782,.779 L.758,.797 L.696,.683 L.662,.692 L.657,.78 L.629,.773 L.599,.801 L.57,.779 L.55,.789 L.515,.753 L.515,.802 L.474,.787 L.447,.826 L.416,.794 L.389,.801 L.353,.729 L.33,.753 L.296,.716 L.242,.756 L.251,.82 L.205,.828 L.175,.804 L.156,.834 L.177,.989 L.128,.997 L.089,.937 L.109,.855 L.054,.805 L.033,.725 L.048,.656 L.02,.619 L0,.431 L.02,.402 L.078,.452 L.07,.5 L.028,.48 L.066,.557 L.061,.521 L.08,.539 L.089,.486 L.103,.497 L.091,.426 L.119,.472 L.156,.408 L.161,.433 L.187,.403 L.193,.418 L.192,.373 L.246,.41 L.221,.333 L.237,.261 L.276,.406 L.272,.442 L.253,.445 L.279,.453 L.283,.377 L.313,.406 L.298,.371 L.272,.369 L.256,.255 L.267,.303 L.292,.31 L.267,.284 L.286,.263 L.346,.347 L.281,.22 L.313,.194 L.317,.215 L.298,.163 L.327,.097 L.364,.092 Z',
  fernA: 'M.584,.024 L.639,.04 L.659,.085 L.564,.283 L.515,.091 L.534,.046 Z M.931,.001 L.998,.031 L.964,.097 L.705,.188 L.595,.288 L.907,.179 L.982,.194 L.986,.249 L.89,.275 L.591,.292 L.521,.364 L.647,.325 L.803,.313 L.887,.331 L.897,.382 L.792,.408 L.512,.374 L.431,.47 L.562,.433 L.689,.428 L.799,.459 L.801,.506 L.694,.529 L.425,.478 L.381,.537 L.674,.564 L.755,.62 L.774,.667 L.72,.684 L.621,.663 L.378,.541 L.349,.593 L.538,.694 L.575,.772 L.524,.801 L.432,.778 L.352,.707 L.321,.628 L.205,.858 L.174,.998 L.186,.872 L.326,.596 L.092,.66 L.026,.647 L.004,.61 L.099,.564 L.316,.587 L.368,.534 L.087,.502 L.013,.46 L.007,.41 L.05,.389 L.118,.403 L.37,.53 L.415,.47 L.149,.354 L.125,.318 L.142,.279 L.206,.255 L.286,.273 L.419,.465 L.519,.348 L.311,.254 L.285,.215 L.298,.179 L.363,.155 L.43,.179 L.526,.341 L.564,.283 L.699,.18 L.809,.034 Z',
  leaf39: 'M.514,0 L.49,.135 L.433,.259 L.597,.269 L.719,.38 L.748,.525 L.676,.676 L.597,.73 L.5,.749 L.403,.73 L.324,.676 L.252,.525 L.262,.424 L.306,.341 L.123,.171 L.033,.321 L0,.499 L.039,.694 L.146,.853 L.305,.961 L.5,1 L.695,.961 L.854,.853 L.961,.694 L1,.499 L.962,.308 L.858,.15 L.703,.043 Z',
  leaf37: 'M.75,.499 L.739,.574 L.693,.658 L.55,.744 L.402,.73 L.28,.619 L.251,.474 L.307,.34 L.425,.261 L.567,.259 L.51,.135 L.485,0 L.296,.043 L.142,.15 L.038,.308 L0,.499 L.039,.694 L.146,.853 L.305,.961 L.499,1 L.6,.99 L.738,.94 L.853,.853 L.94,.738 L.99,.6 L.998,.45 L.963,.311 L.892,.19 L.862,.154 L.685,.332 L.733,.409 Z',
  leaf16: 'M.747,.483 L.742,1 L.849,.981 L.987,.994 L.989,.263 L.945,.049 L.878,0 L.751,.041 L.582,.138 L.131,.478 L.014,.694 L.001,.806 L.02,.874 L.153,.902 L.287,.895 Z',
  leaf27: 'M.96,0 L.81,.12 L.79,.16 L.61,.33 L.49,.49 L.03,.87 L0,.92 L.02,.96 L.09,1 L.14,.99 L.45,.7 L.59,.56 L.85,.21 L.98,.12 L1,.06 Z',
  leaf38: 'M.86,.97 L.96,.79 L1,.6 L.98,.4 L.89,.21 L.74,.07 L.6,.01 L.45,0 L.26,.07 L.09,.25 L.01,.46 L.01,.69 L.09,.91 L.15,.99 L.33,.79 L.25,.63 L.27,.46 L.38,.32 L.53,.29 L.66,.35 L.74,.49 L.74,.66 L.65,.8 L.84,1 Z',
  leaf20: 'M.11,1 L.03,.98 L0,.93 L.04,.57 L.17,.1 L.22,.02 L.3,0 L.4,.07 L.43,.17 L.2,.81 L.46,.85 L.69,.92 L.83,.89 L.94,.94 L1,1 Z',
  leaf15: 'M.11,1 L.01,.97 L.01,.78 L.12,.23 L.18,.04 L.23,0 L.33,.09 L.39,.26 L.2,.8 L.46,.84 L.69,.92 L.83,.89 L.94,.93 L1,1 Z',
  mVein0: 'M1,0 L0,1',
  leaf5: 'M.98,.33 L.96,.21 L.84,.01 L.16,.01 L.02,.26 L.06,.64 L0,1 L1,1 L.94,.64 Z',
  leaf21: 'M.34,.78 L.15,1 L0,.87 L.37,.45 L.66,.03 L.73,0 L.97,.13 L1,.18 L.97,.21 Z',
  leaf22: 'M.96,0 L.79,.12 L.78,.17 L.59,.35 L.47,.52 L.03,.87 L0,.91 L.02,.96 L.08,.99 L.15,1 L.33,.86 L.59,.59 L.85,.22 L.98,.12 L1,.06 Z',
  leaf26: 'M.68,.78 L.9,1 L1,.82 L.36,.13 L.21,.01 L.15,.01 L.03,.17 L.01,.29 L.53,.63 Z',
  leafSpray: 'M.14,0 L.1,.14 L.15,.29 L.48,.42 L.74,.84 L.86,.93 L1,.98 L.85,1 L.67,.94 L.55,.82 L.36,.49 L.05,.36 L0,.27 L.01,.17 L.06,.07 Z',
  leaf6: 'M0,.43 L0,.82 L.77,.98 L.92,1 L1,.97 L.94,.78 L.79,.47 L.62,.19 L.5,.06 L.27,0 L.33,.28 L.36,.36 L.43,.41 L.69,.77 L.09,.72 L.03,.42 Z',
  leaf7: 'M1,.43 L1,.82 L.24,.98 L.08,1 L0,.97 L.06,.78 L.21,.47 L.38,.19 L.5,.06 L.73,0 L.67,.28 L.64,.36 L.57,.41 L.31,.77 L.91,.72 L.97,.42 Z',
  leaf32: 'M.46,0 L.37,.08 L.41,.12 L.49,.08 L.43,.11 L.47,.14 L.43,.16 L.49,.19 L.42,.19 L.44,.25 L.53,.24 L.49,.27 L.53,.27 L.49,.29 L.48,.35 L.47,.28 L.41,.27 L.43,.23 L.33,.18 L.35,.15 L.32,.18 L.32,.12 L.21,.19 L.16,.11 L0,.44 L.02,.44 L.03,.49 L.07,.47 L.09,.61 L.07,.65 L.1,.7 L.08,.74 L.11,.74 L.12,.79 L.48,.83 L.61,.91 L.61,1 L.78,.94 L.81,.89 L.82,.94 L.87,.94 L.84,.98 L.91,.94 L.84,.88 L.87,.86 L.77,.91 L.86,.83 L.94,.84 L1,.8 L.98,.75 L.94,.76 L.98,.74 L.93,.7 L.93,.59 L.86,.65 L.87,.57 L.82,.53 L.77,.52 L.74,.62 L.75,.69 L.7,.72 L.7,.81 L.67,.8 L.67,.7 L.55,.65 L.56,.6 L.53,.61 L.6,.47 L.63,.48 L.6,.44 L.63,.47 L.69,.43 L.65,.39 L.7,.41 L.69,.49 L.72,.46 L.75,.47 L.71,.38 L.77,.38 L.79,.31 L.77,.28 L.83,.29 L.83,.26 L.88,.39 L.85,.45 L.8,.43 L.79,.46 L.84,.47 L.9,.55 L.89,.49 L.93,.54 L.9,.43 L.96,.43 L.96,.48 L1,.44 L.93,.35 L.96,.34 L.94,.33 L.96,.3 L.93,.31 L.96,.29 L.92,.3 L.94,.26 L.89,.25 L.88,.2 L.84,.22 L.85,.15 L.8,.16 L.78,.24 L.79,.18 L.82,.14 L.8,.13 L.73,.21 L.76,.24 L.73,.23 L.74,.26 L.78,.27 L.71,.37 L.72,.31 L.69,.32 L.7,.28 L.68,.27 L.7,.19 L.77,.12 L.72,.1 L.68,.19 L.67,.13 L.7,.11 L.67,.09 L.63,.15 L.65,.2 L.68,.18 L.65,.25 L.67,.28 L.61,.37 L.63,.31 L.56,.32 L.53,.28 L.59,.24 L.57,.18 L.62,.09 L.59,.09 L.59,.11 L.57,.09 L.55,.16 L.55,.1 L.54,.12 L.52,.08 L.49,.1 L.5,.03 Z',
  leaf10: 'M.28,.78 L0,.8 L.03,1 L.3,.99 L.37,.95 L.78,.5 L1,.18 L.74,.01 L.67,.02 L.59,.14 L.56,.3 Z',
  leaf23: 'M.56,.96 L.66,1 L.9,.92 L.99,.74 L.98,.54 L.9,.41 L.77,.29 L.2,0 L0,.3 L.49,.79 Z',
  leaf29: 'M0,0 L.07,.34 L.21,.51 L.75,.16 L.99,.66 L1,.91 L.96,.99 L.8,.74 L.78,.83 L.47,.85 L.23,.93 L.21,1 L.16,.89 L.14,.91 L.06,.84 L.01,.51 L.06,.48 L.02,.44 Z',
  leaf13: 'M.61,.99 L.88,.83 L1,.47 L.59,0 L.28,.23 L0,.54 L.08,.69 L.27,.87 L.49,.99 Z',
  leaf1: 'M.18,.03 L.56,.49 L.88,.74 L.97,.87 L1,1 L.05,.62 L0,.43 L.03,.18 L.1,.01 Z',
  leaf3: 'M.82,.03 L.44,.49 L.12,.74 L.03,.87 L0,1 L.95,.62 L1,.43 L.97,.18 L.9,.01 Z',
  leaf14: 'M.95,1 L.84,.98 L.24,.73 L.23,.68 L.34,.6 L.33,.55 L0,.08 L.83,.01 L.72,.65 L.98,.93 L1,.98 Z',
  mVein2: 'M.98,1 L1,.87 L.96,.76 L.84,.63 L.61,.51 L.55,.35 L.46,.22 L.28,.1 L0,0',
  leaf18: 'M.87,.58 L1,.44 L.96,.23 L.84,.1 L.58,0 L.43,.04 L.28,.14 L.08,.36 L0,.53 L.03,.65 L.15,.8 L.35,.93 L.55,1 L.61,.98 L.71,.81 L.8,.8 Z',
  leaf12: 'M.96,1 L.8,.97 L.2,.74 L.19,.68 L.31,.61 L.31,.55 L0,.07 L.84,.01 L.69,.66 L.79,.79 L.98,.94 L1,.98 Z',
  leaf25: 'M.99,.99 L.64,.83 L.57,.71 L.01,.26 L.01,.17 L.1,.03 L.16,0 L.23,.07 L.69,.62 L.94,.85 Z',
  leaf33: 'M.77,0 L.73,.02 L.74,.13 L.7,.21 L.75,.2 L.77,.25 L.66,.31 L.66,.38 L.58,.43 L.44,.41 L.26,.22 L.17,.36 L.19,.43 L.05,.43 L.07,.48 L0,.53 L.01,.59 L.05,.55 L.07,.6 L.14,.58 L.24,.76 L.38,.83 L.49,.78 L.56,.96 L.6,.98 L.65,.93 L.73,1 L.85,.91 L.9,.76 L.87,.74 L.89,.71 L.84,.69 L.89,.69 L.84,.61 L.88,.53 L.82,.54 L.79,.49 L.85,.42 L.85,.49 L.95,.4 L1,.23 L.99,.16 L.94,.19 Z',
  leaf2: 'M1,.46 L.61,.13 L.32,.01 L.12,.04 L.01,.24 L.01,.48 L.14,.67 L.46,.85 L.86,1 Z',
  leaf4: 'M0,.46 L.39,.13 L.68,.01 L.88,.04 L.99,.24 L.99,.48 L.86,.67 L.54,.85 L.14,1 Z',
  leaf8: 'M1,.46 L.96,.23 L.85,.09 L.69,.02 L.45,0 L.26,.03 L.11,.12 L.02,.28 L0,.5 L.05,.7 L.16,.85 L.33,.96 L.52,1 L.7,.95 L.86,.83 L.97,.66 Z',
  leaf19: 'M.65,.03 L.48,0 L.35,.03 L.17,.15 L.08,.28 L0,.66 L.04,.91 L.1,.99 L.21,.99 L.35,.92 L.5,.69 L.78,.7 L.98,.58 L1,.41 L.95,.27 L.84,.14 Z',
  leaf35: 'M.19,.74 L.16,.75 Z M.68,0 L.82,.04 L.66,.06 L.82,.05 L.8,.07 L.88,.09 L.68,.13 L.86,.13 L.77,.23 L.92,.15 L1,.18 L.87,.22 L.93,.24 L.86,.25 L.91,.26 L.83,.35 L.86,.43 L.79,.44 L.85,.47 L.79,.47 L.84,.51 L.79,.51 L.81,.56 L.67,.56 L.72,.57 L.68,.58 L.76,.63 L.77,.68 L.66,.62 L.7,.65 L.62,.68 L.76,.7 L.61,.76 L.55,.74 L.51,.81 L.37,.84 L.29,1 L.2,.98 L.15,.88 L.18,.85 L.15,.86 L.17,.81 L.13,.81 L.19,.78 L.13,.8 L.13,.77 L.17,.76 L.14,.76 L.16,.72 L.21,.74 L.17,.72 L.24,.66 L.17,.63 L.24,.65 L.24,.6 L.2,.57 L.17,.6 L.2,.55 L.18,.42 L.02,.37 L0,.31 L.09,.31 L0,.25 L.15,.23 L.2,.17 L.16,.15 L.31,.07 L.32,.11 L.32,.07 L.41,.06 L.39,.12 L.41,.09 L.46,.11 L.48,.06 L.55,.14 L.52,.06 L.61,.1 L.56,.03 L.64,.05 Z',
  leaf9: 'M.78,.54 L.83,.97 L1,.89 L.95,.29 L.89,.02 L.72,.02 L.21,.21 L.02,.68 L.03,.94 L.11,1 L.28,.96 Z',
  mVein1: 'M1,1 L.48,.74 L0,.69 L.21,.51 L.31,.29 L.25,.11 L.12,0',
  leaf11: 'M1,.32 L.95,.23 L.61,.04 L.36,0 L.1,.04 L0,.19 L.13,.68 L.08,.84 L.84,1 L.93,.8 L.9,.4 Z',
  leaf31: 'M.85,.78 L.84,.84 Z M.84,.84 L.82,.89 Z M.85,.83 L.83,.91 Z M.82,0 L.54,.12 L.59,.28 L.56,.24 L.56,.29 L.52,.24 L.42,.27 L.44,.36 L.53,.35 L.5,.42 L.34,.51 L.39,.53 L.34,.6 L.4,.57 L.35,.68 L.45,.69 L.4,.76 L.19,.89 L0,.95 L.33,.85 L.48,.74 L.46,.82 L.52,.75 L.48,.77 L.5,.69 L.64,.6 L.54,.71 L.68,.64 L.83,.81 L.81,.93 L.84,1 L.86,.91 L.87,.99 L.86,.76 L.8,.78 L.8,.71 L.76,.71 L1,.2 L.86,.12 Z',
  leaf36: 'M.09,0 L0,.38 L.04,.64 L.09,.71 L.27,.75 L.41,.97 L.47,.83 L.56,.87 L.58,.81 L.65,.83 L.7,1 L.7,.79 L.81,.64 L.79,.53 L.82,.49 L.82,.58 L.86,.44 L.93,.41 L.94,.33 L1,.3 L.99,.2 L.71,.39 L.72,.22 L.55,.08 Z',
  mVein3: 'M.71,1 L.87,.83 L.98,.63 L.99,.45 L.93,.34 L.33,.06 L0,0',
  leaf30: 'M0,.06 L0,0 L.52,.08 L.89,.27 L.99,.23 L1,.31 L.88,.35 L.37,.81 L.34,1 L0,.73 L.06,.2 Z',
  leaf24: 'M.26,.02 L.05,.21 L.02,.57 L.2,.88 L.55,1 L.89,.76 L1,.61 L.56,.22 L.56,.03 Z',
  bit37: 'M.61,.89 L.59,.95 Z M.34,0 L.23,.02 L.27,.07 L.21,.11 L.11,.09 L.1,.23 L.02,.26 L0,.33 L.09,.4 L.22,.37 L.24,.43 L.41,.53 L.43,.69 L.54,.79 L.46,.9 L.57,1 L.67,.78 L.86,.68 L.89,.44 L1,.3 L.87,.2 L.75,.22 L.68,.15 L.62,.22 L.59,.18 L.65,.14 L.56,.16 L.61,.1 L.57,.02 L.52,.09 L.38,.1 Z',
  bit12: 'M.76,1 L.1,.99 L0,.88 L.48,.16 L.63,0 L.84,.08 L.96,.26 L.38,.78 L.98,.86 L.96,.94 Z',
  bit8: 'M.06,.18 L.01,1 L.92,1 L1,.33 L.96,.05 L.62,.01 Z',
  bit28: 'M.78,0 L.73,.21 L.6,.12 L.64,.04 L.52,.01 L.45,.14 L.36,.09 L.23,.26 L.03,.33 L.02,.77 L.43,.67 L.51,.8 L.58,.71 L.53,.82 L.58,.78 L.62,.95 L.71,.94 L.73,1 L.84,.94 L.97,.69 L1,.5 Z',
  bit14: 'M.33,.77 L.44,.94 L.65,1 L.82,.95 L.99,.77 L.9,.38 L.71,.09 L.58,.01 L.27,.06 L0,.33 L.15,.53 L.14,.62 Z',
  bit5: 'M.13,0 L.03,.03 L0,.14 L.12,.18 L.11,.35 L.3,.51 L.15,.8 L.14,.96 L.58,1 L.98,.96 L.91,.52 L.52,.18 L.62,.07 Z',
  bit1: 'M0,0 L.02,.61 L.5,1 L.98,.61 L1,0 Z',
  bit2: 'M1,1 L.95,.42 L.81,.11 L.5,0 L.19,.11 L.05,.42 L0,1 L.57,.73 L.74,.49 L.8,.73 Z',
  bit20: 'M.92,.09 L.58,0 L.27,.06 L0,.29 L.13,.45 L.48,.5 L.69,.71 L.73,.93 L.9,1 L1,.76 L.8,.24 Z',
  bit3: 'M.93,.99 L.28,.96 L0,.07 L.44,.01 L.51,.83 L.97,.89 L1,.97 Z',
  bit36: 'M.57,.88 L.51,.99 Z M.83,.55 L.94,.73 L.99,1 L.88,.81 L.82,.91 Z M.68,.42 L.72,.59 L.77,.48 L.83,.55 L.82,.91 L.77,.68 L.7,.56 L.68,.63 Z M.54,.31 L.53,.38 L.45,.39 L.46,.48 L.51,.44 L.5,.64 L.46,.55 L.44,.71 L.44,.38 Z M.36,.25 L.35,.3 Z M.37,.17 L.36,.24 Z M0,.07 L.15,.37 L.2,.72 L.38,.88 L.49,.88 L.39,.91 L.2,.81 Z M.39,0 L.43,.35 L.35,.63 L.27,.57 L.25,.36 Z',
  bit18: 'M0,.41 L.08,1 L1,1 L1,0 L0,0 Z',
  bit19: 'M1,.41 L.92,1 L0,1 L0,0 L1,0 Z',
  bit15: 'M.86,.8 L1,.53 L.85,.21 L.47,.04 L.04,.05 L.03,.31 L.2,.47 L.05,.72 L.09,.94 L.45,.99 Z',
  sparkle: 'M.51,0 L.55,.44 L1,.48 L.55,.51 L.5,1 L.44,.51 L0,.48 L.44,.44 Z',
  bit6: 'M.66,.22 L.56,.05 L.34,0 L.17,.05 L.01,.24 L.1,.63 L.43,.99 L.73,.93 L1,.66 L.86,.37 Z',
  bit30: 'M.25,0 L.03,.25 L.09,.37 L0,.4 L.1,.52 L.17,.45 L.2,.69 L.37,1 L.45,.92 L.45,.71 L.74,.47 L.7,.27 L.84,.33 L.87,.47 L1,.21 L.95,.14 L.8,.26 L.69,.21 L.68,.27 L.59,.27 L.39,.19 L.4,.12 Z',
  bit29: 'M.5,0 L.31,.15 L.34,.37 L.19,.4 L.08,.32 L0,.47 L.06,.67 L.17,.64 L.11,.81 L.26,1 L.26,.76 L.33,.73 L.49,.82 L.55,.99 L.71,.84 L.87,.88 L.85,.71 L.98,.53 L.98,.35 L.8,.3 L.7,.05 L.62,.13 Z',
  bit4: 'M0,.82 L.29,.31 L.4,.4 L.94,0 L.96,.95 L.36,.9 L.1,1 Z',
  bit11: 'M.02,.87 L.04,.54 L.27,.21 L.61,.01 L.92,.08 L.99,.38 L.74,.75 L.26,1 L.11,.98 Z',
  bit9: 'M.22,.86 L.79,.95 L1,.18 L.71,.17 L.37,0 L.28,.51 L0,1 Z',
  bit24: 'M.08,0 L0,0 L.04,.29 L.19,.56 L.07,.04 L.36,.76 L.79,1 L.84,.84 L.95,.82 L.99,.64 L.72,.82 L.62,.64 L.65,.4 L.53,.2 L.46,.24 L.36,.07 Z',
  bit25: 'M.21,0 L.06,.09 L0,.29 L.06,.37 L0,.55 L.11,.8 L.06,.92 L.4,1 L.32,.93 L.48,.83 L.33,.76 L.54,.65 L.42,.59 L.85,.5 L.89,.44 L.75,.37 L.77,.25 L1,.14 L.7,.16 L.71,.09 Z',
  bit23: 'M.62,0 L.57,.1 L.46,.04 L.45,.15 L.39,.08 L.27,.14 L.35,.18 L.29,.25 L.41,.24 L.28,.25 L.3,.33 L.44,.29 L.34,.36 L.59,.31 L.51,.38 L.57,.4 L.38,.41 L.39,.51 L.3,.38 L.25,.58 L.32,.63 L.19,.6 L.14,.69 L.23,.65 L.18,.79 L.26,.76 L.18,.81 L.11,.72 L.13,.82 L0,.87 L.09,.96 L.19,.89 L.28,.99 L.32,.91 L.22,.85 L.44,.78 L.49,.68 L.42,.62 L.52,.63 L.44,.59 L.61,.59 L.86,.39 L.72,.39 L.84,.35 L.8,.31 L.99,.28 L.98,.14 Z',
  bit21: 'M.05,.16 L.05,0 L1,.15 L.29,.57 L.4,.86 L.63,1 L.08,.91 L.16,.81 L0,.62 L.1,.5 Z',
  bit10: 'M.02,.86 L.04,.53 L.31,.16 L.65,0 L.94,.11 L.98,.43 L.74,.75 L.26,1 Z',
  bit7: 'M.13,.21 L0,.48 L.15,.8 L.54,.97 L.97,.95 L.96,.68 L.79,.53 L.94,.28 L.89,.06 L.53,.01 Z',
  bit27: 'M.84,0 L.5,.16 L.24,.64 L.01,.71 L.12,.78 L0,.79 L.09,.83 L.04,.96 L.15,.99 L.25,.87 L.3,.94 L.3,.6 L.5,.24 L.99,.13 Z',
  bit26: 'M.29,0 L.29,.24 L.1,.18 L.01,.41 L.33,.93 L.62,1 L1,.38 L.94,.28 L.86,.33 L.87,.1 L.62,.22 Z',
  bit22: 'M.19,0 L.1,.02 L.16,.39 L0,.48 L.1,.53 L.06,.65 L.16,.6 L.19,.69 L.01,.74 L.35,.93 L1,.98 L.33,.83 L.42,.73 L.24,.54 L.34,.39 L.23,.32 L.5,.12 Z',
  bit13: 'M.18,.66 L.69,1 L1,.55 L.53,0 L0,.65 Z',
  bit34: 'M.72,0 L.32,.11 L.37,.29 L0,.49 L.6,1 L1,.77 L.87,.61 L.82,.01 Z',
  bit16: 'M1,0 L0,0 L.01,.69 L.11,.87 L.5,1 L.89,.87 L.99,.69 Z',
  bit32: 'M.2,0 L0,.26 L.41,.96 L.46,.89 L.59,1 L.67,.86 L.99,.75 L.99,.58 L.83,.57 L.65,.22 Z',
  bit35: 'M.7,0 L.36,.02 L.21,.48 L0,.6 L.24,.6 L.28,.71 L.5,.67 L.53,.88 L.92,1 L.89,.41 L1,.1 Z',
  bit31: 'M.02,0 L.25,.65 L.54,.89 L.93,.99 L1,.88 L.87,.69 L.83,.21 L.61,.09 L.38,.22 L.18,.01 Z',
  bit33: 'M.89,0 L.17,.01 L0,.51 L.37,.94 L.64,1 L.82,.9 L.68,.76 L1,.22 Z',
  bit17: 'M1,.5 L.85,.85 L.5,1 L.15,.85 L0,.5 L.15,.15 L.5,0 L.85,.15 Z',
  icoFb: 'M1,.5 L.85,.85 L.5,1 L.15,.85 L0,.5 L.15,.15 L.5,0 L.85,.15 Z',
  icoCircle: 'M1,.5 L.85,.85 L.5,1 L.15,.85 L0,.5 L.15,.15 L.5,0 L.85,.15 Z',
  checkMark: 'M.64,.37 L.31,.52 L.47,.62 Z M.5,0 L.85,.15 L1,.5 L.85,.85 L.5,1 L.15,.85 L0,.5 L.15,.15 Z',
  icoWa: 'M.34,.27 L.44,.52 L.75,.59 L.72,.69 L.38,.58 L.27,.39 Z M.47,.08 L.12,.34 L.12,.88 L.63,.88 L.89,.65 L.85,.26 Z M.49,0 L.93,.24 L.95,.71 L.65,.95 L0,1 L.06,.29 Z',
  icoInsta: 'M.5,.33 L.33,.5 L.5,.67 L.67,.5 Z M.5,.24 L.76,.5 L.5,.76 L.24,.5 Z M.77,.17 L.77,.29 Z M.5,.09 L.11,.21 L.1,.74 L.2,.89 L.74,.9 L.89,.8 L.88,.19 Z M.5,0 L.96,.14 L.98,.81 L.78,.99 L.19,.98 L.01,.78 L.02,.19 L.17,.03 Z',
  icoTw: 'M0,.89 L.31,1 L.67,.86 L.97,.02 L.6,.03 L.49,.31 L.07,.05 L.05,.44 L.3,.78 Z',
  chevron: 'M.04,.12 L.7,.5 L.02,.89 L.05,.99 L1,.52 L.2,.01 L.03,.03 L.04,.12',
  icoFbGlyph: 'M.23,1 L.66,1 L.66,.5 L.97,.5 L1,.33 L.66,.33 L.67,.21 L1,.18 L1,0 L.37,.04 L.23,.33 L0,.33 L0,.5 L.23,.5 Z',
};

function outline(d, w, h) {
  return d.split(' ').map(tok => {
    if (tok === 'Z') return { close: true };
    const [x, y] = tok.slice(1).split(',');
    return { x: +x * w, y: +y * h, moveTo: tok[0] === 'M' };
  });
}

/* Draw a list of shapes. Each op is [type, x, y, w, h, opts?] with inches;
   `type` is a pptxgenjs preset name or a key of ART.
   `f` is the fill colour (0 = none) and the trailing opts object holds
   a=fill transparency, ln=line colour, lw=line width, fh/fv=flip,
   r=rotation, rr=corner radius.                                           */
function draw(s, ops) {
  for (const [t, x, y, w, h, f, o] of ops) {
    const q = o || {};
    const opt = { x, y, w, h, line: { type: 'none' } };
    if (f) opt.fill = q.a ? { color: f, transparency: q.a } : { color: f };
    if (q.ln) opt.line = { color: q.ln, width: q.lw || 1 };
    if (q.fh) opt.flipH = true;
    if (q.fv) opt.flipV = true;
    if (q.r) opt.rotate = q.r;
    if (q.rr !== undefined) opt.rectRadius = q.rr;
    if (ART[t]) { opt.points = outline(ART[t], w, h); s.addShape('custGeom', opt); }
    else s.addShape(t, opt);
  }
}

/* Repeat one shape at a list of [x, y] positions. */
function rep(s, type, w, h, fill, pts, o) {
  draw(s, pts.map(([x, y]) => [type, x, y, w, h, fill, o]));
}

/* Text box. `body` is a string, or [text, opts] runs; br:1 ends a line.
   opts: sz=pt, b/i/u, c=colour, al=align (c/r/j), lh=line spacing,
         sa=space after paragraph (pt), va=middle, nw=no wrap.             */
function T(s, body, x, y, w, h, o) {
  const q = o || {};
  const runs = (typeof body === 'string' ? [[body]] : body).map(([t, r]) => {
    const ro = r || {};
    const has = k => ro[k] !== undefined;
    return {
      text: t,
      options: {
        fontSize: has('sz') ? ro.sz : q.sz, bold: !!(has('b') ? ro.b : q.b),
        italic: !!(has('i') ? ro.i : q.i), underline: (has('u') ? ro.u : q.u) ? { style: 'sng' } : undefined,
        color: has('c') ? ro.c : q.c, breakLine: !!ro.br,
      },
    };
  });
  s.addText(runs, {
    x, y, w, h, fontFace: FONT, fontSize: q.sz || 18, color: q.c || INK,
    align: { c: 'center', r: 'right', j: 'justify' }[q.al] || 'left',
    valign: q.va ? 'middle' : 'top', wrap: !q.nw, paraSpaceAfter: q.sa,
    lineSpacingMultiple: q.lh, margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true,
  });
}

/* A monstera leaf: filled blade plus its four white veins. */
function monstera(s, x, y, w, h, fill, o) {
  const q = o || {};
  const v = (k, dx, dy, dw, dh) => [k, x + dx * w, y + dy * h, dw * w, dh * h, 0,
    { ln: W, lw: 1.75, fh: q.fh, r: q.r }];
  draw(s, [
    ['monstera', x, y, w, h, fill, { fh: q.fh, r: q.r }],
    v('mVein0', 0.095, 0.388, 0.723, 0.464), v('mVein1', 0.674, 0.188, 0.225, 0.423),
    v('mVein2', 0.421, 0.249, 0.242, 0.568), v('mVein3', 0.223, 0.512, 0.201, 0.401),
  ]);
}

/* The four social badges that sit in the footer of several slides. */
function social(s, x, y) {
  const glyph = ['icoInsta', 'icoFbGlyph', 'icoTw', 'icoWa'];
  const gx = [0.072, 0.486, 0.844, 1.224], gy = [0.071, 0.068, 0.082, 0.069];
  const gw = [0.127, 0.062, 0.129, 0.133], gh = [0.127, 0.134, 0.105, 0.133];
  const ops = [];
  for (let i = 0; i < 4; i++) {
    ops.push([i === 1 ? 'icoFb' : 'icoCircle', x + i * 0.3843, y, 0.27, 0.27, ORNG]);
    ops.push([glyph[i], x + gx[i], y + gy[i], gw[i], gh[i], CREAM]);
  }
  draw(s, ops);
}

/* A five-star rating row. */
function stars(s, x, y, size, step) {
  draw(s, [0, 1, 2, 3, 4].map(i => ['star5', x + i * step, y, size, size, ORNG]));
}

/* Repeat one text style at a list of [body, x, y] positions. */
function repT(s, w, h, o, items) {
  for (const [body, x, y] of items) T(s, body, x, y, w, h, o);
}

/* Twenty-four rounded dots orbiting a centre - the circular gauges. */
function ring(s, cx, cy, radius, dot, fill) {
  const ops = [];
  for (let i = 0; i < 24; i++) {
    const a = (i / 24) * 2 * Math.PI;
    ops.push(['roundRect', cx + radius * Math.cos(a) - dot / 2, cy + radius * Math.sin(a) - dot / 2,
      dot, dot, fill, { rr: dot / 2.4, r: (i / 24) * 360 }]);
  }
  draw(s, ops);
}

/* Stand-in for the product photographs (phone / laptop / watch mock-ups) of
   the original deck: a soft panel of the same box, with a caption. */
function photo(s, x, y, w, h, o) {
  const q = o || {};
  draw(s, [['roundRect', x, y, w, h, SAND2, { rr: 0.14, fh: q.fh }]]);
  T(s, '[image]', x, y + h / 2 - 0.2, w, 0.4, { sz: 12, c: SAGE2, al: 'c' });
}

/* The orange pill button ("Discover More" etc.) used across the deck. */
function button(s, x, y, label, fill) {
  const i = label.indexOf(' ');
  draw(s, [['roundRect', x, y, 1.7, 0.47, fill || GOLD, { rr: 0.06 }]]);
  T(s, [[label.slice(0, i), { b: 1 }], [label.slice(i)]], x, y, 1.7, 0.47,
    { sz: 14, c: W, al: 'c', va: 1 });
}

/* Bottom-left brand tag and bottom-right publish note. */
function brandTag(s, y) {
  T(s, [['HEALTH_YOGA_FOUNDATION_'], ['YOGAES', { b: 1 }]], 0.25, y, 1.76, 0.68,
    { sz: 12, c: ORNG, lh: 1.5 });
}
function published(s, y, c) {
  T(s, [['PUBLISHED : '], ['NOW', { b: 1 }]], 11.32, y, 1.76, 0.38,
    { sz: 12, c: c || ORNG, al: 'r', lh: 1.5 });
}

function slide01(s) {
  draw(s, [
    ['blobBig',4.14,4.02,4.38,3.49,SAND], ['blobCorner',11.02,0,2.31,1.65,CREAM2,{fh:1}],
    ['blobCorner',-0,5.93,2.19,1.56,CREAM2,{fh:1,r:180}], ['fernB',4.3,3.67,1.75,2.43,ORNG2,{fh:1,r:338.6}],
    ['fernB',6.79,4.26,1.26,1.75,ORNG2,{fh:1,r:42.8}], ['sparkle',8.86,3.96,.4,.67,ORNG,{fh:1}],
    ['sparkle',3.99,3.1,.4,.67,ORNG,{fh:1}], ['bit1',5.7,6.67,1.11,.34,'291701'],
    ['leaf1',5.68,6.79,1.49,.66,SAND4], ['leaf2',4.82,6.47,1.12,.74,BROWN], ['leaf3',5.34,6.79,1.49,.66,SAND4],
    ['leaf4',6.57,6.47,1.12,.74,BROWN], ['leaf5',5.7,4.65,1.11,2.02,ORNG], ['leaf6',6.27,4.67,1.39,1.25,SAND4],
    ['leaf7',4.86,4.67,1.38,1.25,SAND4], ['leaf8',5.91,3.45,.7,1.07,SAND4], ['bit2',5.9,3.46,.72,.51,BROWN],
    ['fernA',7.99,5.45,1.35,2.05,ORNG,{fh:1,fv:1,r:180}], ['fernA',3.82,6.02,.97,1.48,ORNG,{fh:1}],
    ['leafSpray',10.5,5.32,1.75,1.64,SAND,{r:90}], ['leafSpray',.3,1.06,1.75,1.64,SAND,{r:184.1}],
  ]);
  monstera(s,2.98,4.74,2.87,2.37,CREAM2);
  monstera(s,6.84,5.08,2.31,1.91,CREAM2);
  social(s,11.55,6.91);
  T(s,'YOGAES',2.54,.58,7.93,2.42,{sz:138,b:1,c:GOLD,al:'c'});
  brandTag(s,6.5);
  T(s,L.healthyYogaFound,4.46,2.39,3.93,.44,{sz:20,c:ORNG,al:'c'});
  published(s,6.46);
}

function slide02(s) {
  draw(s, [
    ['blobSoft',11.3,5.48,1.8,2.26,SAND,{fh:1,r:270}], ['blobWave',0,0,1.74,1.23,SAND],
    ['sparkle',6.09,5.38,.4,.67,ORNG,{fh:1}], ['sparkle',1.01,2.19,.4,.67,ORNG,{fh:1}],
    ['fernB',-.22,.35,1.26,1.75,ORNG2,{fh:1,r:42.8}],
  ]);
  monstera(s,10.86,5.74,2.12,1.75,CREAM2);
  social(s,8.26,5.99);
  T(s,'“Today, Modern Physics Is Proving To You That As You Sit Here, Every Subatomic Particle In Your Body Is In Constant Transaction With Everything Else In The Existence. If This Transaction Stops, You Will Cease To Exist. So, Yoga Means To Know The Union Of Existence By Experience”',5.92,1.8,6.09,3.09,{sz:20,i:1,c:ORNG,al:'c',lh:1.5});
  T(s,'Yoga & Health Quote',7.29,1.06,3.36,.51,{sz:24,c:ORNG,al:'c'});
  button(s,8.12,5.22,'Discover More');
}

function slide03(s) {
  draw(s, [
    ['roundRect',8.54,2.11,3.6,4.14,W,{rr:.1}], ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}],
    ['blobCorner',0,4.96,3.58,2.56,CREAM2,{fh:1,r:180}], ['fernA',8.42,6.02,.97,1.48,ORNG,{fh:1}],
    ['roundRect',8.58,1.58,3.47,1.05,W], ['roundRect',8.66,1.66,3.31,.89,SAND2],
  ]);
  monstera(s,11.14,5.3,2.06,1.71,CREAM2);
  stars(s,9.63,5.48,.232,.303);
  T(s,'Welcome To Yogaes Health & Yoga Foundations',1.42,1,6.38,1.31,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf3,1.42,3.18,6.22,2.19,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'Beautiful Body And Peacefull Mind Built On You',1.42,2.68,5.1,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.allAspectsOf,1.42,5.53,6.22,.78,{sz:14,c:ORNG,al:'j',lh:1.5,sa:6});
  T(s,'Drop Blood Preasure',8.96,1.72,2.73,.37,{sz:16,b:1,c:ORNG2,al:'c'});
  T(s,L.benefitOfYoga,9.37,2.12,1.9,.37,{sz:16,c:SAGE2,al:'c'});
}

function slide04(s) {
  draw(s, [
    ['blobWave',0,0,1.74,1.23,SAND], ['fernA',4.21,.51,1.35,2.05,ORNG,{fh:1,fv:1,r:180}],
    ['fernB',.85,1.08,1.25,1.74,ORNG2,{fh:1,r:338.6}], ['sparkle',5.64,4.24,.4,.67,ORNG,{fh:1}],
    ['sparkle',1.09,.65,.4,.67,ORNG,{fh:1}], ['blobSoft',11.3,5.48,1.8,2.26,SAND,{fh:1,r:270}],
    ['checkMark',6.42,2.62,.26,.26,SAND3], ['checkMark',6.42,4.28,.26,.26,SAND3],
    ['blobSide',.67,1.31,5.08,2.65,'FFF7EC'], ['blobSide2',.94,1.45,4.54,2.36,SAND],
  ]);
  monstera(s,4.27,1.59,1.88,1.56,CREAM2);
  monstera(s,.46,2.51,1.13,.94,CREAM2);
  T(s,L.allAspectsOf3,6.84,4.19,5.58,1.89,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.todayModernPhysi,6.84,2.46,5.58,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'Keep Healthy With Yoga',6.42,1.51,5.87,.71,{sz:36,b:1,c:ORNG});
  T(s,'Peacefull',1.37,1.83,3.69,1.01,{sz:54,b:1,i:1,c:SAGE2,al:'c'});
  T(s,'Minds',1.72,2.46,2.77,1.01,{sz:54,b:1,i:1,c:SAGE2,al:'c'});
}

function slide05(s) {
  draw(s, [
    ['blobBig',.29,-0,9.77,7.13,SAND,{fv:1}], ['roundRect',9.53,.7,2.77,3.19,W,{rr:.227}],
    ['roundRect',8.53,4.12,2.37,2.73,W,{rr:.195}], ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}],
    ['fernA',.05,.1,.97,1.48,ORNG,{fh:1,r:195}],
  ]);
  monstera(s,11.16,5.36,2.06,1.71,CREAM2);
  rep(s,'ellipse',.36,.36,ORNG,[[2.45,4.54],[2.45,5.25],[5,4.54],[5,5.25]]);
  rep(s,'chevron',.08,.14,W,[[2.6,4.65],[2.6,5.36],[5.16,4.65],[5.16,5.36]],{ln:W});
  repT(s,1.77,.37,{sz:16,b:1,c:SAGE2},[[L.balanceBody,2.9,4.57],['Peacefull Mind',2.9,5.25],
    [L.healthyLife,5.46,4.57]]);
  T(s,L.todayModernPhysi,2.45,2.35,5.58,1.84,{sz:14,c:SAGE2,al:'j',lh:1.5,sa:6});
  T(s,L.startFocusOn,2.45,.94,5.48,1.31,{sz:36,b:1,c:SAGE2});
  T(s,'Drop Blood Preasure',5.46,5.25,2.48,.37,{sz:16,b:1,c:SAGE2});
}

function slide06(s) {
  draw(s, [
    ['blobWave',0,0,3.98,2.81,CREAM2], ['roundRect',.63,.69,12.07,4.12,W,{rr:.207}],
    ['blobSoft',10.81,4.97,2.24,2.81,SAND,{fh:1,r:270}], ['fernB',11.83,3.64,1.26,1.75,ORNG2,{fh:1,r:42.8}],
    ['fernA',.36,4.07,.97,1.48,ORNG,{fh:1}], ['ellipse',5.02,3,.88,.88,SAND,{a:75,ln:SAND}],
    ['ellipse',5.12,3.09,.69,.69,SAND], ['ellipse',5.02,5,.88,.88,SAND,{a:75,ln:SAND}],
    ['ellipse',5.12,5.09,.69,.69,SAND],
  ]);
  T(s,'01',5.12,3.09,.69,.69,{sz:14,c:SAGE2,al:'c',va:1});
  T(s,'Everyday Doing A Full Meditations For Better Feel',5.07,1.4,6.99,1.31,{sz:36,b:1,c:ORNG});
  T(s,'Peacefull Mind',6.04,3,1.77,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.allAspectsOf2,6.04,3.44,5.68,.98,{sz:12,c:INK,al:'j',lh:1.5,sa:6});
  T(s,L.theMoreYou,6.04,5.37,5.68,.98,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'02',5.12,5.09,.69,.69,{sz:14,c:SAGE2,al:'c',va:1});
  T(s,L.balanceBody,6.04,5,1.77,.37,{sz:16,b:1,c:ORNG2});
}

function slide07(s) {
  draw(s, [
    ['blobBig',4.43,1.07,8.8,6.43,SAND], ['ellipse',4.45,1.78,3.45,3.97,W],
    ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}], ['fernA',3.94,3.76,.97,1.48,ORNG,{fh:1}],
    ['sparkle',3.83,2.58,.4,.67,ORNG,{fh:1}], ['sparkle',7.26,.81,.4,.67,ORNG,{fh:1}],
    ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}], ['checkMark',.39,2.35,.26,.26,SAND3],
  ]);
  monstera(s,11.08,.92,2.06,1.71,CREAM2);
  T(s,'100% Cerified Healthy',8.01,2.24,3.85,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.todayModernPhysi,8.19,3.64,3.47,2.2,{sz:12,c:INK,al:'c',lh:1.5,sa:6});
  button(s,9.08,6.02,'Discover More');
  T(s,L.todayModernPhysi,.81,3.04,2.86,2.8,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'Beautiful Body And Better Mind Built On You',.81,2.3,2.86,.64,{sz:16,b:1,c:ORNG2});
}

function slide08(s) {
  draw(s, [
    ['fernB',3.15,6.02,1.25,1.74,ORNG2,{r:21.4}], ['blobSoft',.41,3.8,3.31,4.13,CREAM2,{r:90}],
    ['fernA',1.5,4.95,1.35,2.05,ORNG,{fv:1,r:180}], ['bit3',3.86,6.54,.61,.86,SAND],
    ['leaf9',2.49,5.96,1.65,.67,BROWN], ['bit4',.28,7.12,1.25,.29,SAND], ['leaf10',1.43,6.02,1.62,1.38,BROWN],
    ['bit5',2.61,3.17,.3,2.08,SAND4], ['leaf11',2.5,4.76,.65,1.51,ORNG], ['bit6',2.3,4.03,.67,.56,'F4B382'],
    ['bit7',2.24,4.01,.43,.61,BROWN], ['bit5',2.54,3.1,.3,2.08,SAND],
    ['blobWave',11.59,0,1.74,1.23,SAND,{fh:1}], ['sparkle',1.19,2.95,.4,.67,ORNG],
    ['sparkle',5.33,6.08,.4,.67,ORNG], ['ellipse',5.17,2.48,.59,.59,ORNG2],
    ['ellipse',5.17,3.88,.59,.59,ORNG2],
  ]);
  monstera(s,2.6,4.7,2.06,1.71,CREAM2);
  T(s,'01',5.17,2.48,.59,.59,{sz:11,c:SAGE2,al:'c',va:1});
  T(s,L.allAspectsOf2,5.9,2.28,5.68,.98,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'02',5.17,3.88,.59,.59,{sz:11,c:SAGE2,al:'c',va:1});
  T(s,L.allAspectsOf2,5.9,3.68,5.68,.98,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'24 Hours Ready To Register All New Members',5.17,.73,6.24,1.31,{sz:36,b:1,c:ORNG});
  button(s,10.73,5.08,'Learns More');
  T(s,L.allAspectsOf,5.19,4.92,5.26,.78,{sz:14,c:ORNG,al:'j',lh:1.5,sa:6});
}

function slide09(s) {
  draw(s, [
    ['blobSoft',11.55,-.2,1.58,1.98,SAND,{fh:1,fv:1,r:90}], ['blobCorner',9.75,4.96,3.58,2.56,CREAM2,{r:180}],
    ['roundRect',3.94,1.02,7,1.97,W], ['roundRect',4.1,1.16,6.68,1.67,SAND2], ['roundRect',3.94,4.37,7,1.97,W],
    ['roundRect',4.1,4.52,6.68,1.67,SAND2],
  ]);
  monstera(s,11.25,.54,2.06,1.71,CREAM2);
  stars(s,4.52,2.2,.232,.302);
  stars(s,4.52,5.62,.232,.302);
  social(s,11.55,6.91);
  T(s,L.balanceBody,4.4,4.77,2.52,.4,{sz:18,b:1,c:ORNG2});
  T(s,L.benefitOfYoga,4.4,5.14,1.9,.37,{sz:16,c:SAGE2});
  T(s,'Drop Blood Preasure',4.4,1.35,2.73,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.benefitOfYoga,4.4,1.75,1.9,.37,{sz:16,c:SAGE2});
  T(s,L.todayModernPhysi3,6.91,1.31,2.55,1.34,{sz:10,c:INK,lh:1.5,sa:6});
  T(s,L.todayModernPhysi3,6.91,4.77,2.55,1.34,{sz:10,c:INK,lh:1.5,sa:6});
  brandTag(s,6.5);
  published(s,6.46);
}

function slide10(s) {
  draw(s, [
    ['blobWave',9.12,4.53,4.21,2.97,CREAM2,{fh:1,fv:1}], ['fernB',8.29,4.93,2.09,2.91,ORNG2,{r:293}],
    ['leaf12',11.3,2.87,.71,2.09,SAND3], ['bit8',11.29,1.94,.65,1.09,'333333'],
    ['leaf13',9.22,2.1,1.17,1.38,BROWN], ['leaf14',11.33,2.87,.72,2.12,SAND],
    ['leaf15',9.66,4.82,1.98,1.87,SAND3], ['leaf16',9.55,1.31,2.41,1.73,BROWN],
    ['leaf17',8.92,2.82,1.69,2.95,ORNG], ['bit9',10.35,5.26,.72,.43,SAND3],
    ['leaf18',10.77,4.58,1.07,1.4,SAND], ['leaf19',10.68,4.51,1.16,1,BROWN],
    ['bit10',10.73,4.46,.64,.42,'4C230C'], ['bit11',10.68,4.4,.69,.46,BROWN],
    ['leaf20',9.63,4.95,1.98,1.92,SAND], ['blobSoft',.23,-.23,1.8,2.26,SAND,{fv:1,r:270}],
    ['sparkle',7.55,2.01,.4,.67,ORNG,{fh:1}], ['sparkle',12.51,3.33,.4,.67,ORNG,{fh:1}],
    ['fernA',11.83,5.45,1.35,2.05,ORNG,{fh:1,fv:1,r:180}], ['leafSpray',11.09,.42,1.75,1.64,SAND,{r:184.1}],
  ]);
  monstera(s,.62,.39,1.7,1.41,CREAM2);
  T(s,L.startFocusOn,1.43,2.2,5.82,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.allAspectsOf,1.22,3.69,6.22,.78,{sz:14,c:ORNG,al:'c',lh:1.5,sa:6});
  T(s,L.theMoreYou,1.49,4.67,5.68,.98,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6});
  button(s,3.48,6,'Discover More');
}

function slide11(s) {
  draw(s, [
    ['blobBig',-.24,.72,6.55,6.06,SAND,{fv:1,r:270}], ['roundRect',1.66,.84,3.93,3.05,W,{rr:.25}],
    ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}], ['fernA',.05,.1,.97,1.48,ORNG,{fh:1,r:195}],
    ['checkMark',6.25,1.64,.26,.26,SAND3], ['checkMark',6.25,2.51,.26,.26,SAND3],
    ['checkMark',1.74,4.67,.26,.26,SAGE],
  ]);
  monstera(s,11.16,5.46,2.06,1.71,CREAM2);
  T(s,'Peacefull Mind',6.25,1.04,1.77,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.youWillCease,6.67,1.48,4.4,.78,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.todayModernPhysi4,6.67,2.35,5.25,1.13,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.todayModernPhysi,2.15,4.51,7.08,1.49,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  button(s,7.54,6.22,'Learns More');
}

function slide12(s) {
  draw(s, [
    ['fernB',4.05,3.36,1.26,1.75,ORNG2,{fh:1,r:42.8}], ['sparkle',2.33,2.8,.4,.67,ORNG,{fh:1}],
    ['fernA',.97,5.02,.97,1.48,ORNG,{fh:1}], ['roundRect',6.83,.46,5.79,6.52,SAND,{rr:.292}],
    ['bit12',1.01,6.18,.79,.96,SAND], ['leaf21',1.47,4.61,1.96,1.84,BROWN],
    ['leaf22',4.33,3.09,1.6,2.17,SAND3], ['leaf23',2.87,4.5,1.93,1.06,ORNG], ['bit13',4.49,5.25,.56,.38,SAND3],
    ['bit14',4.61,5.4,.85,.73,SAND], ['bit15',5,5.37,.54,.79,BROWN], ['leaf24',2.7,4.48,1,.85,BROWN],
    ['leaf25',4.68,6.03,1.36,1.08,SAND], ['leaf26',2.89,4.7,2.03,1.64,BROWN],
    ['leaf27',4.35,3.07,1.74,2.21,SAND], ['blobCorner',-0,5.93,2.19,1.56,CREAM2,{fh:1,r:180}],
    ['fernA',11.74,5.45,1.35,2.05,ORNG,{fh:1,fv:1,r:180}], ['sparkle',6.04,5.81,.4,.67,ORNG,{fh:1}],
  ]);
  monstera(s,.54,3.65,2.87,2.37,CREAM2);
  monstera(s,4.27,4.15,2.31,1.91,CREAM2);
  rep(s,'ellipse',.36,.36,ORNG,[[7.46,.73],[7.46,2.24],[7.46,3.75],[7.46,5.38]]);
  rep(s,'chevron',.08,.14,W,[[7.61,.84],[7.61,2.35],[7.61,3.86],[7.61,5.49]],{ln:W});
  T(s,L.startFocusOn,.71,.51,5.82,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.theMoreYou,.78,1.94,5.68,.98,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6});
  repT(s,1.77,.37,{sz:16,b:1,c:SAGE2},[[L.balanceBody,7.91,.76],['Peacefull Mind',7.91,2.25],
    [L.healthyLife,7.91,3.77]]);
  repT(s,4.58,.84,{sz:10,c:SAGE2,al:'j',lh:1.5,sa:6},[[L.todayModernPhysi2,7.45,1.25],
    [L.todayModernPhysi2,7.45,2.72],[L.todayModernPhysi2,7.45,4.23],[L.todayModernPhysi2,7.45,5.86]]);
  T(s,'Drop Blood Preasure',7.91,5.38,2.48,.37,{sz:16,b:1,c:SAGE2});
}

function slide13(s) {
  draw(s, [
    ['blobBig',4.43,1.07,8.8,6.43,SAND], ['roundRect',1.31,.69,5.18,6.43,W,{rr:.261}],
    ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}], ['fernA',.76,5.82,.97,1.48,ORNG,{fh:1}],
    ['sparkle',11.74,.48,.4,.67,ORNG,{fh:1}], ['sparkle',7.26,.81,.4,.67,ORNG,{fh:1}],
    ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  rep(s,'ellipse',1.69,1.94,SAND,[[1.74,.99],[4.19,.99],[1.74,3.92],[4.19,3.92]]);
  monstera(s,11.47,2.4,2.06,1.71,CREAM2);
  T(s,'Qualified Personal For Teach You About Yoga',6.94,2.44,5.33,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.youWillCease,7.4,3.83,4.4,.78,{sz:14,c:INK,al:'c',lh:1.5,sa:6});
  repT(s,1.58,.38,{sz:14,c:SAND,al:'c',lh:1.2},[[L.expertStaff,1.75,3.46],[L.expertStaff,4.2,3.46],
    [L.expertStaff,1.75,6.39],[L.expertStaff,4.2,6.39]]);
  repT(s,1.64,.46,{sz:18,b:1,c:SAGE,al:'c',lh:1.2},[['Yuli Veria',1.75,3.05],['Celicia Tuie',1.75,5.98],
    ['Barrett Kunz',4.2,5.98]]);
  T(s,'Jerica Whited',3.98,3.05,2.08,.46,{sz:18,b:1,c:SAGE,al:'c',lh:1.2});
  T(s,L.todayModernPhysi,7.3,4.8,4.6,1.59,{sz:12,c:INK,al:'c',lh:1.5,sa:6});
}

function slide14(s) {
  draw(s, [
    ['blob1',6.78,-.97,5.59,7.52,SAND,{fh:1,fv:1,r:270}], ['ellipse',7.44,.82,5.12,5.88,W],
    ['fernA',-.04,5.96,.97,1.48,ORNG,{fh:1,r:15}], ['sparkle',12.16,1.39,.4,.67,ORNG,{fh:1}],
    ['sparkle',8.2,.37,.4,.67,ORNG,{fh:1}], ['fernB',12.43,.38,1.1,1.52,ORNG2,{r:317.2}],
    ['roundRect',1.23,3.92,.6,.09,SAGE,{rr:.029}], ['ellipse',1.25,5.7,.21,.21,'B5CC2A'],
    ['ellipse',1.21,5.66,.28,.28,ORNG,{ln:'B5CC2A'}], ['ellipse',1.16,5.62,.37,.37,0,{ln:'B5CC2A'}],
    ['plus',1.27,5.72,.16,.16,W], ['roundRect',4.26,4.92,1.17,.09,SAGE,{rr:.029}],
    ['roundRect',4.27,3.92,1.51,.09,SAGE,{rr:.029}], ['blobSide',7.59,4.58,5.08,2.65,'FFF7EC'],
    ['blobSide2',7.86,4.72,4.54,2.36,SAND], ['ellipse',1.25,6.25,.21,.21,'B5CC2A'],
    ['ellipse',1.21,6.21,.28,.28,ORNG,{ln:'B5CC2A'}], ['ellipse',1.16,6.17,.37,.37,0,{ln:'B5CC2A'}],
    ['plus',1.27,6.27,.16,.16,W],
  ]);
  rep(s,'roundRect',1.86,.09,'A6A6A6',[[1.23,3.92],[4.26,4.92],[4.27,3.92]],{a:85,rr:.029});
  stars(s,3.74,5.72,.207,.27);
  monstera(s,11.19,4.86,1.88,1.56,CREAM2);
  monstera(s,7.37,5.78,1.13,.94,CREAM2);
  stars(s,3.74,6.27,.207,.27);
  T(s,'Gloria Dorces',1.2,.96,3.37,.71,{sz:36,b:1,c:ORNG});
  T(s,'All aspects of physicality in the universe are cyclical. Planets are going around the sun, the solar system is moving, and everything in the galaxy and the cosmos is cyclical. The more you are identified with your physical system, the more cyclical you are. ',1.2,1.7,5.39,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  repT(s,2.71,.34,{sz:14,c:ORNG},[['Communication Management',1.17,3.45],['Yoga Knowledge',4.2,4.45],
    ['Professional Trainer',4.2,3.45]]);
  T(s,'35%',3.18,3.82,.7,.28,{sz:10.5,c:ORNG});
  T(s,[['High Education',{br:1}],['Best Skill Medical',{br:1}],
    ['5+ Award Medical']],1.17,4.24,3.82,1.03,{sz:12,c:ORNG,lh:1.3});
  T(s,L.professionalYoga,1.54,5.68,2.06,.3,{sz:12,c:CREAM,nw:1});
  T(s,'50%',6.21,4.82,.7,.29,{sz:10.5,c:ORNG});
  T(s,'85%',6.21,3.82,.7,.29,{sz:10.5,c:ORNG});
  T(s,'Gloria',8.29,5.1,3.69,1.01,{sz:54,b:1,i:1,c:SAGE2,al:'c'});
  T(s,'Dorces',8.63,5.73,2.77,1.01,{sz:54,b:1,i:1,c:SAGE2,al:'c'});
  T(s,L.professionalYoga,1.54,6.24,2.06,.3,{sz:12,c:CREAM,nw:1});
}

function slide15(s) {
  draw(s, [
    ['blobSoft',11.55,-.2,1.58,1.98,SAND,{fh:1,fv:1,r:90}], ['blobCorner',9.75,4.96,3.58,2.56,CREAM2,{r:180}],
    ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}], ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  monstera(s,11.25,.54,2.06,1.71,CREAM2);
  rep(s,'ellipse',.36,.36,ORNG,[[10.33,2.6],[10.33,3.42],[10.33,4.3]]);
  rep(s,'chevron',.08,.14,W,[[10.48,2.71],[10.48,3.53],[10.48,4.41]],{ln:W});
  T(s,'Self Control',6.67,2.52,1.76,.28,{sz:10.5,c:ORNG});
  repT(s,.55,.28,{sz:10.5,c:ORNG,al:'r'},[['60%',9.37,2.52],['93%',9.37,3.05],['55%',9.38,3.75],
    ['65%',9.38,4.3]]);
  repT(s,2.54,.28,{sz:10.5,c:ORNG},[['Mentality',6.67,3.05],['Communication',6.67,3.75],
    [' Professional Trainer',6.67,4.3]]);
  repT(s,1.27,.71,{sz:36,b:1,c:ORNG},[['25+',1.16,5.57],['120',4.38,5.57],['99%',2.83,5.57],['890',6.06,5.57]]);
  T(s,'Awards',1.17,6.24,1.37,.4,{sz:18,b:1,c:CREAM});
  T(s,'Achivement',4.4,6.23,1.66,.4,{sz:18,b:1,c:CREAM});
  T(s,'Skill',2.85,6.24,1.37,.4,{sz:18,b:1,c:CREAM});
  T(s,'Experience',6.07,6.23,1.66,.4,{sz:18,b:1,c:CREAM});
  T(s,'My Name, Julia Yuicva',6.8,1.44,5.46,.71,{sz:36,b:1,c:ORNG});
  T(s,L.theMoreYou,7.75,5.27,4.69,1.29,{sz:12,c:CREAM,lh:1.5,sa:6});
  repT(s,1.77,.37,{sz:16,b:1,c:CREAM},[[L.balanceBody,10.78,2.63],['Peacefull Mind',10.78,3.42],
    [L.healthyLife,10.78,4.33]]);
}

function slide16(s) {
  draw(s, [
    ['blobSoft',11.55,-.2,1.58,1.98,SAND,{fh:1,fv:1,r:90}], ['blobCorner',9.75,4.96,3.58,2.56,CREAM2,{r:180}],
    ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}], ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  monstera(s,11.25,.54,2.06,1.71,CREAM2);
  rep(s,'blobSide',3.54,1.85,'FFF7EC',[[4.99,2.64],[8.83,2.47],[.9,2.47]]);
  rep(s,'blobSide2',3.17,1.65,SAND,[[5.17,2.73],[9.02,2.57],[1.08,2.57]]);
  monstera(s,7.5,2.83,1.31,1.09,CREAM2);
  monstera(s,4.84,3.47,.79,.65,CREAM2);
  monstera(s,11.35,2.67,1.31,1.09,CREAM2);
  monstera(s,8.69,3.31,.79,.65,CREAM2);
  monstera(s,3.41,2.67,1.31,1.09,CREAM2);
  monstera(s,.75,3.31,.79,.65,CREAM2);
  T(s,L.qualifiedService,3.48,1.04,6.38,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  repT(s,1.27,.71,{sz:36,b:1,c:SAGE2,al:'c'},[['$10',2.01,2.85],['25+',6.14,3.09],['4M',9.86,2.69]]);
  T(s,'Peacefull Mind',1.66,3.52,1.96,.4,{sz:18,b:1,c:SAGE2,al:'c'});
  T(s,L.healthyLife,5.47,3.76,2.62,.4,{sz:18,b:1,c:SAGE2,al:'c'});
  T(s,'Drop Blood Preasure',9.21,3.36,2.57,.4,{sz:18,b:1,c:SAGE2,al:'c'});
  repT(s,3.06,2.54,{sz:14,c:CREAM,al:'c',lh:1.5,sa:6},[[L.theMoreYou,1.11,4.35],[L.theMoreYou,8.96,4.35],
    [L.theMoreYou,5.13,4.54]]);
}

function slide17(s) {
  draw(s, [
    ['blobWave',7.14,4.53,4.21,2.97,CREAM2,{fh:1,fv:1}], ['fernB',7.1,4.93,2.09,2.91,ORNG2,{r:293}],
    ['blobSoft',.23,-.23,1.8,2.26,SAND,{fv:1,r:270}], ['sparkle',7.33,2.46,.4,.67,ORNG,{fh:1}],
    ['sparkle',10.05,.38,.4,.67,ORNG,{fh:1}],
  ]);
  monstera(s,.62,.39,1.7,1.41,CREAM2);
  T(s,'All Guaranted Services For Your Meditation',.72,1.76,5.46,1.31,{sz:36,b:1,c:ORNG});
  repT(s,1.39,.71,{sz:36,b:1,c:CREAM,al:'r'},[['01.',.13,3.27],['02.',.13,4.36],['03.',.09,5.48]]);
  repT(s,1.58,.38,{sz:14,c:CREAM,lh:1.2},[[L.expertService,1.52,3.65],[L.expertService,1.52,4.74],
    [L.expertService,1.48,5.86]]);
  T(s,L.balanceBody,1.52,3.24,2.09,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.meditationTime,1.52,4.33,2.32,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.peacefulMinds,1.48,5.45,2.09,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  repT(s,3.6,.68,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6},[[L.youWillCease,3.84,3.26],[L.youWillCease,3.84,4.36],
    [L.youWillCease,3.84,5.48]]);
}

function slide18(s) {
  draw(s, [
    ['fernB',3.15,6.02,1.25,1.74,ORNG2,{r:21.4}], ['blobSoft',.41,3.8,3.31,4.13,CREAM2,{r:90}],
    ['fernA',1.5,4.95,1.35,2.05,ORNG,{fv:1,r:180}], ['blobWave',11.59,0,1.74,1.23,SAND,{fh:1}],
    ['sparkle',.8,.74,.4,.67,ORNG], ['sparkle',5.33,6.08,.4,.67,ORNG],
  ]);
  monstera(s,10.56,.22,2.06,1.71,CREAM2);
  rep(s,'ellipse',.59,.59,'FACC6B',[[10.8,1.09],[10.8,3.11],[10.8,5.16]]);
  repT(s,.59,.59,{sz:11,c:SAGE2,al:'c',va:1},[['01',10.8,1.09],['02',10.8,3.11],['03',10.8,5.16]]);
  repT(s,1.58,.38,{sz:14,c:CREAM,al:'r',lh:1.2},[[L.expertService,9.07,1.39],[L.expertService,8.99,3.41],
    [L.expertService,8.96,5.46]]);
  T(s,L.balanceBody,8.74,.97,1.91,.46,{sz:18,b:1,c:CREAM,al:'r',lh:1.2});
  T(s,L.meditationTime,8.33,3,2.24,.46,{sz:18,b:1,c:CREAM,al:'r',lh:1.2});
  T(s,L.peacefulMinds,8.3,5.04,2.24,.46,{sz:18,b:1,c:CREAM,al:'r',lh:1.2});
  T(s,'Routine Yoga Services For Your Meditation',1.15,1.4,5.52,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  repT(s,3.42,.68,{sz:12,c:CREAM,al:'r',lh:1.5,sa:6},[[L.youWillCease,7.97,1.82],[L.youWillCease,7.97,3.79],
    [L.youWillCease,7.97,5.92]]);
  T(s,L.allAspectsOf,1.07,2.91,5.68,.78,{sz:14,c:ORNG,al:'c',lh:1.5,sa:6});
  T(s,L.theMoreYou,1.07,3.9,5.68,.98,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6});
  button(s,3.06,5.22,'Discover More');
}

function slide19(s) {
  draw(s, [
    ['blobSoft',10.81,4.97,2.24,2.81,SAND,{fh:1,r:270}], ['leaf12',2.83,4.73,.25,.73,SAND3],
    ['leaf13',2.1,4.46,.41,.48,BROWN], ['leaf14',2.84,4.73,.25,.74,SAND], ['leaf15',2.25,5.41,.69,.65,SAND3],
    ['leaf16',2.21,4.19,.84,.6,BROWN], ['leaf17',1.99,4.71,.59,1.03,ORNG], ['leaf18',2.64,5.33,.37,.49,SAND],
    ['leaf19',2.61,5.3,.41,.35,BROWN], ['leaf20',2.24,5.46,.69,.67,SAND], ['bit12',4.98,5.65,.4,.48,SAND],
    ['leaf21',5.21,4.86,.98,.93,BROWN], ['leaf22',6.65,4.1,.8,1.09,SAND3], ['leaf23',5.91,4.81,.97,.53,ORNG],
    ['bit14',6.79,5.26,.43,.37,SAND], ['leaf24',5.83,4.8,.5,.43,BROWN], ['leaf25',6.82,5.58,.68,.54,SAND],
    ['leaf26',5.92,4.91,1.02,.82,BROWN], ['leaf27',6.66,4.09,.87,1.11,SAND], ['bit3',11.29,5.61,.34,.48,SAND],
    ['leaf9',10.52,5.28,.93,.38,BROWN], ['leaf10',9.92,5.32,.91,.78,BROWN], ['bit5',10.59,3.71,.17,1.17,SAND4],
    ['leaf11',10.53,4.61,.36,.85,ORNG], ['bit5',10.55,3.67,.17,1.17,SAND],
    ['blobSoft',11.55,-.2,1.58,1.98,SAND,{fh:1,fv:1,r:90}],
    ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}], ['fernB',-.27,.54,1.1,1.52,ORNG2,{fh:1,r:42.8}],
    ['fernA',.7,5.48,.78,1.18,ORNG,{fh:1}],
  ]);
  rep(s,'round2DiagRect',3.05,3.44,SAND,[[1.29,2.99],[4.99,2.99],[9.19,2.99]]);
  rep(s,'round2DiagRect',3.18,3.58,W,[[1.01,2.69],[4.7,2.69],[8.9,2.69]]);
  monstera(s,11.25,.54,2.06,1.71,CREAM2);
  monstera(s,8.46,5.3,1.08,.9,CREAM2);
  T(s,L.qualifiedService,3.36,.81,6.38,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,'Peacefull Mind',1.66,3.02,1.96,.4,{sz:18,b:1,c:SAGE2,al:'c'});
  T(s,L.healthyLife,5.1,3.02,2.62,.4,{sz:18,b:1,c:SAGE2,al:'c'});
  T(s,'Drop Blood Preasure',9.21,3.02,2.57,.4,{sz:18,b:1,c:SAGE2,al:'c'});
  repT(s,2.8,.84,{sz:10,c:INK,al:'c',lh:1.5,sa:6},[[L.youWillCease,1.24,3.44],[L.youWillCease,4.92,3.44],
    [L.youWillCease,9.09,3.44]]);
  button(s,1.66,6.22,'Learns More');
  button(s,5.47,6.22,'Learns More');
  button(s,9.64,6.22,'Learns More');
}

function slide20(s) {
  draw(s, [
    ['blobCorner',11.02,0,2.31,1.65,CREAM2,{fh:1}], ['blobCorner',-0,5.93,2.19,1.56,CREAM2,{fh:1,r:180}],
    ['blobBig',6.64,3.08,5.55,4.42,SAND], ['fernB',6.85,2.65,2.22,3.08,ORNG2,{fh:1,r:338.6}],
    ['fernB',10.01,3.39,1.6,2.22,ORNG2,{fh:1,r:42.8}], ['sparkle',12.62,3.02,.5,.85,ORNG,{fh:1}],
    ['sparkle',6.45,1.93,.5,.85,ORNG,{fh:1}], ['bit1',8.62,6.45,1.41,.43,'291701'],
    ['leaf1',8.61,6.6,1.89,.83,SAND4], ['leaf2',7.51,6.19,1.43,.94,BROWN], ['leaf3',8.16,6.6,1.89,.83,SAND4],
    ['leaf4',9.72,6.19,1.43,.94,BROWN], ['leaf5',8.62,3.89,1.41,2.56,ORNG],
    ['leaf6',9.34,3.91,1.76,1.58,SAND4], ['leaf7',7.56,3.91,1.75,1.58,SAND4],
    ['bit16',9.12,3.59,.43,.43,SAND3], ['leaf8',8.89,2.37,.89,1.36,SAND4], ['bit17',9.1,2.24,.47,.32,'291701'],
    ['bit2',8.87,2.37,.92,.65,BROWN], ['fernA',11.53,4.91,1.71,2.59,ORNG,{fh:1,fv:1,r:180}],
    ['fernA',6.24,5.62,1.24,1.87,ORNG,{fh:1}], ['leafSpray',10.49,.93,1.75,1.64,SAND,{r:90}],
    ['leafSpray',.46,5.31,1.75,1.64,SAND,{r:184.1}],
  ]);
  monstera(s,5.18,4.01,3.63,3.01,CREAM2);
  monstera(s,10.07,4.43,2.92,2.42,CREAM2);
  social(s,11.55,6.91);
  T(s,'5 Min',.25,1.56,6.07,2.42,{sz:138,b:1,c:GOLD,al:'c'});
  brandTag(s,6.5);
  T(s,'Breaks Yoga For A While',1.25,3.36,3.93,.44,{sz:20,c:ORNG,al:'c'});
  published(s,6.46,CREAM);
}

function slide21(s) {
  draw(s, [
    ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}], ['blobCorner',0,4.96,3.58,2.56,CREAM2,{fh:1,r:180}],
    ['fernA',.88,5.92,.97,1.48,ORNG,{fh:1}],
  ]);
  monstera(s,11.14,5.3,2.06,1.71,CREAM2);
  rep(s,'roundRect',2.24,3.13,W,[[6.75,5.02],[9.15,3.94],[11.39,2.57],[5.52,1.69],[7.92,.61],[10.16,-.76]],{r:345,rr:.185});
  T(s,'The Galery Of Yogaes Health & Yoga Foundations',.66,1.16,4.56,1.92,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.allAspectsOf,.78,3.27,4.3,1.13,{sz:14,c:ORNG,al:'c',lh:1.5,sa:6});
  T(s,L.theMoreYou,.78,4.6,4.3,1.59,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6});
}

function slide22(s) {
  draw(s, [
    ['blobSoft',.23,-.23,1.8,2.26,SAND,{fv:1,r:270}], ['roundRect',.98,.43,11.38,3.16,W],
    ['roundRect',1.23,.67,10.88,2.69,SAND2], ['blobWave',7.14,4.53,4.21,2.97,CREAM2,{fh:1,fv:1}],
    ['fernB',8.55,4.8,2.09,2.91,ORNG2,{r:293}], ['fernB',.86,.42,1.07,1.49,ORNG2,{fh:1,r:338.6}],
    ['fernA',11.29,.81,.93,1.41,ORNG,{fh:1}], ['ellipse',4.66,3.83,4.8,3.4,W],
  ]);
  monstera(s,.6,3.12,1.7,1.41,CREAM2);
  rep(s,'roundRect',2.24,.81,W,[[1.73,6.15],[7.52,6],[10.41,3.3]],{rr:.041});
  monstera(s,8.08,3.41,1.31,1.09,CREAM2);
  T(s,'The Galery of Yoga Meditation With Us',3.81,.79,5.71,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  repT(s,1.58,.38,{sz:14,c:SAGE2,al:'c',lh:1.2},[[L.yogaProgram,2.06,6.15],[L.yogaProgram,7.85,6],
    [L.yogaProgram,10.74,3.3]]);
  T(s,L.balanceBody,1.9,6.41,1.88,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2});
  T(s,'Meditations',7.82,6.26,1.64,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2});
  T(s,L.peacefulMind,10.59,3.56,1.88,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2});
  T(s,L.theMoreYou,3.09,2.11,7.14,.98,{sz:12,c:INK,al:'c',lh:1.5,sa:6});
}

function slide23(s) {
  draw(s, [
    ['blobWave',0,0,3.98,2.81,CREAM2], ['blobSoft',10.81,4.97,2.24,2.81,SAND,{fh:1,r:270}],
    ['fernB',-.27,.54,1.1,1.52,ORNG2,{fh:1,r:42.8}], ['fernA',12.36,6.32,.78,1.18,ORNG,{fh:1}],
  ]);
  rep(s,'roundRect',2.15,.5,W,[[10.31,3.44],[10.31,6.17],[7.62,3.44],[7.62,6.17]],{rr:.124});
  T(s,'Get Best Experience Of Yoga Meditation With Us',.9,2.01,6.18,1.31,{sz:36,b:1,c:ORNG});
  repT(s,2.22,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2},[[L.balanceBody,10.28,6.2],['Meditations',7.58,3.46],
    [L.peacefulMind,10.28,3.46],[L.mentalHealth,7.58,6.18]]);
  T(s,L.allAspectsOf3,.9,3.44,6.22,2.19,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.allAspectsOf,.9,5.79,6.22,.78,{sz:14,c:ORNG,al:'j',lh:1.5,sa:6});
}

function slide24(s) {
  draw(s, [
    ['blobBig',.24,3.96,4.85,3.54,SAND,{fh:1}], ['blobCorner',10.83,0,2.48,1.78,CREAM2,{fv:1,r:180}],
    ['fernA',4.45,6.04,.97,1.48,ORNG], ['sparkle',5.17,4.66,.4,.67,ORNG], ['sparkle',1.34,2.12,.4,.67,ORNG],
    ['fernB',12.48,.38,1.1,1.52,ORNG2,{r:317.2}],
  ]);
  monstera(s,.09,3.64,2.06,1.71,CREAM2);
  rep(s,'ellipse',3.45,3.97,W,[[1.49,2.45],[5.56,2.45],[9.47,2.45]]);
  rep(s,'ellipse',.88,.88,SAND,[[1.85,2.3],[5.92,2.38],[9.84,1.99]],{a:75,ln:SAND});
  rep(s,'ellipse',.69,.69,SAND,[[1.95,2.4],[6.01,2.48],[9.94,2.08]]);
  T(s,L.weHaveSpecial,2.6,.41,8.13,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  repT(s,.69,.69,{sz:14,c:W,al:'c',va:1},[['01',1.95,2.4],['02',6.01,2.48],['03',9.94,2.08]]);
  T(s,L.yogaProgram,6.52,4.37,1.58,.38,{sz:14,c:SAGE2,al:'c',lh:1.2});
  T(s,L.peacefulMind,6.24,4.63,2.13,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2});
  T(s,'Yoga Meditation Program For Healthy Lifes',6.22,5.15,2.17,.68,{sz:12,c:INK,al:'c',lh:1.5,sa:6});
  T(s,L.yogaProgram,10.4,4.37,1.58,.38,{sz:14,c:SAGE2,al:'c',lh:1.2});
  T(s,L.mentalHealth,10.13,4.63,2.13,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2});
  T(s,'Yoga Meditation Program For Healthy Lifes',10.11,5.15,2.17,.68,{sz:12,c:INK,al:'c',lh:1.5,sa:6});
}

function slide25(s) {
  draw(s, [
    ['blobBig',7.79,1.41,6.4,4.67,SAND,{r:270}], ['roundRect',7.12,1,4.28,5.79,W,{rr:.216}],
    ['blobWave',0,0,2.11,1.49,CREAM2], ['fernB',-.27,.54,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  rep(s,'roundRect',2.16,.68,W,[[7.8,1.08],[7.8,3.09],[7.81,5.04]],{rr:.054});
  T(s,'Our Vision Is To Give The Best Mental Health All Members',1.04,1.09,4.74,2.52,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,'BALANCE BODY',8.48,1.24,1.38,.34,{sz:12,b:1,c:ORNG,lh:1.2});
  repT(s,.52,.37,{sz:16,b:1,c:SAGE2,al:'r'},[['01',7.93,1.24],['02',7.93,3.25],['03',7.93,5.2]]);
  T(s,'MEDITATIONS',8.48,3.25,1.38,.34,{sz:12,b:1,c:ORNG,lh:1.2});
  T(s,'PEACEFUL MIND',8.49,5.2,1.6,.34,{sz:12,b:1,c:ORNG,lh:1.2});
  repT(s,2.51,.84,{sz:10,c:INK,lh:1.5,sa:6},[[L.youWillCease,8.48,1.93],[L.youWillCease,8.48,4.06],
    [L.youWillCease,8.48,5.87]]);
  T(s,L.allAspectsOf,1.04,3.84,4.3,1.13,{sz:14,c:ORNG,al:'c',lh:1.5,sa:6});
  T(s,L.theMoreYou,.85,5.17,4.68,1.29,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6});
}

function slide26(s) {
  draw(s, [
    ['blobSoft',10.81,4.97,2.24,2.81,SAND,{fh:1,r:270}], ['round2DiagRect',1.02,1.42,3.69,5.39,SAND],
    ['round2DiagRect',.68,.95,3.83,5.61,W], ['blobSoft',11.55,-.2,1.58,1.98,SAND,{fh:1,fv:1,r:90}],
    ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}], ['fernB',-.27,.54,1.1,1.52,ORNG2,{fh:1,r:42.8}],
    ['fernA',.85,5.96,.78,1.18,ORNG,{fh:1}],
  ]);
  monstera(s,11.25,.54,2.06,1.71,CREAM2);
  monstera(s,3.9,.91,1.08,.9,CREAM2);
  rep(s,'ellipse',.59,.59,ORNG,[[5.02,3.87],[5.02,4.82],[5.02,5.76]]);
  repT(s,.59,.59,{sz:11,c:W,al:'c',va:1},[['01',5.02,3.87],['02',5.02,4.82],['03',5.02,5.76]]);
  repT(s,1.58,.38,{sz:14,c:CREAM,lh:1.2},[[L.yogaProgram,5.77,3.87],[L.yogaProgram,5.77,4.82],
    [L.yogaProgram,5.77,5.76]]);
  T(s,'Meditations',5.73,4.12,1.64,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.weHaveSpecial,.92,1.71,3.06,3.74,{sz:36,b:1,c:ORNG});
  repT(s,3.27,.58,{sz:10,c:CREAM,lh:1.5,sa:6},[[L.youWillCease,7.88,3.87],[L.youWillCease,7.88,4.82],
    [L.youWillCease,7.88,5.76]]);
  T(s,L.peacefulMind,5.73,5.07,1.83,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.balanceBody,5.73,6.02,1.83,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  button(s,2.19,5.55,'Discover More');
}

function slide27(s) {
  draw(s, [
    ['blobSoft',11.3,5.48,1.8,2.26,SAND,{fh:1,r:270}], ['blobWave',0,0,1.74,1.23,SAND],
    ['fernB',-.22,.35,1.26,1.75,ORNG2,{fh:1,r:42.8}],
  ]);
  photo(s,7.68,.86,5.09,6.64,{fh:1});
  monstera(s,10.86,5.74,2.12,1.75,CREAM2);
  T(s,L.weHaveSpecial2,.84,1.73,6.11,1.31,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf3,.84,3.79,6.22,2.19,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'Beautiful Body And Peacefull Mind Built On You',.84,3.29,5.1,.37,{sz:16,b:1,c:ORNG2});
}

function slide28(s) {
  draw(s, [
    ['blobBig',0,-0,7.09,7.13,SAND,{fv:1}], ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}],
    ['fernA',.05,.1,.97,1.48,ORNG,{fh:1,r:195}],
  ]);
  monstera(s,11.16,5.36,2.06,1.71,CREAM2);
  photo(s,4.67,1.29,3.26,5.21);
  rep(s,'ellipse',.59,.59,ORNG,[[8.6,2.48],[8.6,3.43],[8.6,4.38]]);
  T(s,'The Mockup of Yoga Meditation With Us',.81,1.65,4.1,1.92,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.theMoreYou,.91,3.75,3.9,1.59,{sz:12,c:INK,al:'c',lh:1.5,sa:6});
  button(s,2.01,5.78,'Discover More');
  repT(s,.59,.59,{sz:11,c:W,al:'c',va:1},[['01',8.6,2.48],['02',8.6,3.43],['03',8.6,4.38]]);
  repT(s,1.58,.38,{sz:14,c:CREAM,lh:1.2},[[L.yogaProgram,9.34,2.48],[L.yogaProgram,9.34,3.43],
    [L.yogaProgram,9.34,4.38]]);
  T(s,'Meditations',9.31,2.74,1.64,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.peacefulMind,9.31,3.69,1.83,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.balanceBody,9.31,4.64,1.83,.46,{sz:18,b:1,c:ORNG,lh:1.2});
}

function slide29(s) {
  draw(s, [
    ['fernA',10.95,1.22,1.35,2.05,ORNG,{fh:1,fv:1,r:180}], ['fernB',6.71,1.33,1.25,1.74,ORNG2,{fh:1,r:338.6}],
    ['bit18',6.43,5.81,3.21,.14,'B3B4B5'], ['bit19',9.6,5.81,3.21,.14,'B3B4B5'],
    ['blob2',7.06,2.28,5.17,3.54,'D2D3D5'], ['blob3',7.08,2.3,5.13,3.5,'D9D9D9'],
    ['rect',6.43,5.75,6.38,.12,'F2F2F2'], ['rect',7.25,2.52,4.79,3.03,'F2F2F2'],
    ['ellipse',9.61,2.39,.05,.06,'2C2C2C'], ['ellipse',9.61,2.39,.05,.05,'0A0A0A'],
    ['ellipse',9.62,2.39,.03,.04,'000000'], ['ellipse',9.63,2.4,.02,.02,'2C99B4'],
    ['blobWave',0,0,1.74,1.23,SAND], ['sparkle',12.28,3.76,.4,.67,ORNG,{fh:1}],
    ['sparkle',1.09,.65,.4,.67,ORNG,{fh:1}], ['blobSoft',11.3,5.48,1.8,2.26,SAND,{fh:1,r:270}],
    ['checkMark',.51,3.19,.26,.26,SAND3], ['checkMark',.51,4.84,.26,.26,SAND3],
  ]);
  T(s,'We Have Special Mockup For All Member Yoga',.51,1.55,5.91,1.31,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf3,.93,4.75,5.58,1.89,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.todayModernPhysi,.93,3.02,5.58,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
}

function slide30(s) {
  draw(s, [
    ['blobBig',7.05,1.07,6.19,6.43,SAND], ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}],
    ['fernA',6.77,5.21,.97,1.48,ORNG,{fh:1}], ['sparkle',12.19,3.66,.4,.67,ORNG,{fh:1}],
    ['sparkle',7.26,.81,.4,.67,ORNG,{fh:1}], ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}],
    ['ellipse',.87,2.92,.88,.88,SAND,{a:75,ln:SAND}], ['ellipse',.97,3.01,.69,.69,SAND],
    ['ellipse',.87,4.92,.88,.88,SAND,{a:75,ln:SAND}], ['ellipse',.97,5.01,.69,.69,SAND],
  ]);
  monstera(s,11.08,.92,2.06,1.71,CREAM2);
  photo(s,7.44,1.29,2.88,3.27);
  photo(s,8.88,3.09,2.88,3.27,{fh:1});
  T(s,'Qualified Mockup For All Member of Yoga',.87,1.17,5.35,1.31,{sz:36,b:1,c:ORNG});
  T(s,'01',.97,3.01,.69,.69,{sz:14,c:SAGE2,al:'c',va:1});
  T(s,'Peacefull Mind',1.89,2.92,1.77,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.allAspectsOf2,1.89,3.36,4.78,.98,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.theMoreYou,1.89,5.29,4.78,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'02',.97,5.01,.69,.69,{sz:14,c:SAGE2,al:'c',va:1});
  T(s,L.balanceBody,1.89,4.92,1.77,.37,{sz:16,b:1,c:ORNG2});
}

function slide31(s) {
  draw(s, [
    ['blobBig',-.24,.72,6.55,6.06,SAND,{fv:1,r:270}], ['roundRect',5.21,3.39,7.46,2.54,W,{rr:.208}],
    ['blob4',3.24,2.67,3.11,3.79,'D4D4D4'], ['blob5',3.35,2.6,3.02,3.69,'D4D4D4'],
    ['blob6',3.36,2.62,3,3.67,'525252'], ['blob7',3.37,2.62,2.99,3.66,'1F1F1F'],
    ['blob8',1.23,1.29,2.04,4.39,'D4D4D4'], ['blob9',1.07,1.27,2,4.36,'F0F0F0'],
    ['blob10',1.08,1.29,1.99,4.33,'525252'], ['blob11',1.08,1.29,1.98,4.32,'1F1F1F'],
    ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}], ['fernA',.05,.1,.97,1.48,ORNG,{fh:1,r:195}],
    ['ellipse',5.78,3.8,.59,.59,ORNG2], ['ellipse',5.78,4.92,.59,.59,ORNG2],
  ]);
  monstera(s,11.16,5.46,2.06,1.71,CREAM2);
  T(s,'24 Hours Ready To Register New Members',6.08,1.52,5.7,1.31,{sz:36,b:1,c:ORNG});
  T(s,'01',5.78,3.8,.59,.59,{sz:11,c:SAGE2,al:'c',va:1});
  T(s,L.allAspectsOf2,6.51,3.61,5.68,.98,{sz:12,c:SAGE2,al:'j',lh:1.5,sa:6});
  T(s,'02',5.78,4.92,.59,.59,{sz:11,c:SAGE2,al:'c',va:1});
  T(s,L.allAspectsOf2,6.51,4.72,5.68,.98,{sz:12,c:SAGE2,al:'j',lh:1.5,sa:6});
  button(s,10.24,5.75,'Learns More');
}

function slide32(s) {
  draw(s, [
    ['fernB',3.13,5.67,1.25,1.74,ORNG2,{r:21.4}], ['blobSoft',.41,3.8,3.31,4.13,CREAM2,{r:90}],
    ['fernA',.39,3.95,1.35,2.05,ORNG,{fv:1,r:180}], ['blobWave',11.59,0,1.74,1.23,SAND,{fh:1}],
    ['sparkle',.8,.74,.4,.67,ORNG], ['sparkle',5.33,6.08,.4,.67,ORNG],
  ]);
  monstera(s,10.56,.22,2.06,1.71,CREAM2);
  photo(s,.6,1.73,7.28,4.85,{fh:1});
  rep(s,'ellipse',.36,.36,ORNG,[[6.17,2.52],[6.17,3.23],[8.73,2.52],[8.73,3.23]]);
  rep(s,'chevron',.08,.14,W,[[6.32,2.63],[6.32,3.34],[8.88,2.63],[8.88,3.34]],{ln:W});
  T(s,'Our Mission Is To Give The Best Yoga Meditations',6.17,.79,6.18,1.31,{sz:36,b:1,c:ORNG});
  repT(s,1.77,.37,{sz:16,b:1,c:CREAM},[[L.balanceBody,6.62,2.54],['Peacefull Mind',6.62,3.23],
    [L.healthyLife,9.18,2.54]]);
  T(s,'Drop Blood Preasure',9.18,3.23,2.48,.37,{sz:16,b:1,c:CREAM});
  T(s,L.todayModernPhysi,6.12,3.94,5.58,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
}

function slide33(s) {
  draw(s, [
    ['blobBig',6.12,2.31,7.12,5.2,SAND], ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}],
    ['fernA',6.47,3.98,.97,1.48,ORNG,{fh:1}], ['sparkle',11.74,.48,.4,.67,ORNG,{fh:1}],
    ['sparkle',7.26,.81,.4,.67,ORNG,{fh:1}], ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}],
    ['blobSide',6.45,4.46,5.08,2.65,'FFF7EC'], ['blobSide2',6.72,4.6,4.54,2.36,SAND],
  ]);
  monstera(s,11.47,2.4,2.06,1.71,CREAM2);
  photo(s,7.16,1.54,6.17,5.2);
  monstera(s,10.05,4.75,1.88,1.56,CREAM2);
  monstera(s,6.24,5.66,1.13,.94,CREAM2);
  T(s,L.weHaveSpecial2,.6,1.49,6.11,1.31,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.allAspectsOf,.55,2.98,6.22,.78,{sz:14,c:ORNG,al:'c',lh:1.5,sa:6});
  T(s,L.theMoreYou,.81,3.97,5.68,.98,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6});
  button(s,2.81,5.29,'Discover More');
  T(s,'Healthy',7.15,4.98,3.69,1.01,{sz:54,b:1,i:1,c:SAGE2,al:'c'});
  T(s,'Lifes',7.5,5.61,2.77,1.01,{sz:54,b:1,i:1,c:SAGE2,al:'c'});
}

function slide34(s) {
  draw(s, [
    ['fernB',8.93,6.09,1.25,1.74,ORNG2,{fh:1,r:338.6}], ['blobSoft',9.62,3.8,3.31,4.13,CREAM2,{fh:1,r:270}],
    ['fernA',10.48,5.03,1.35,2.05,ORNG,{fh:1,fv:1,r:180}], ['blobWave',0,0,1.74,1.23,SAND],
    ['sparkle',12.13,.74,.4,.67,ORNG,{fh:1}], ['sparkle',7.61,6.08,.4,.67,ORNG,{fh:1}],
    ['roundRect',2.33,2.37,2.83,3.45,W,{rr:.143}], ['roundRect',8.13,2.37,2.83,3.45,W,{rr:.143}],
    ['roundRect',5.07,2.04,3.1,4.03,ORNG2,{rr:.156}],
  ]);
  monstera(s,.71,.3,2.06,1.71,CREAM2);
  social(s,11.55,6.91);
  T(s,'$18,10',2.65,4.64,2.25,.57,{sz:28,b:1,c:ORNG,al:'c'});
  T(s,L.theFeatures,2.73,3.24,2.11,.3,{sz:12,c:SAGE,al:'c'});
  T(s,L.youWillCease,2.58,3.49,2.41,.91,{sz:11,c:INK,al:'c',lh:1.5});
  T(s,L.getTheDiscount,2.73,5.23,2.11,.3,{sz:12,i:1,c:SAGE,al:'c'});
  T(s,'BASIC',3.03,2.6,1.5,.51,{sz:24,b:1,c:ORNG,al:'c'});
  T(s,L.chooseYourFavori,2.37,1.04,8.59,.71,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,'$98,10',8.46,4.64,2.25,.57,{sz:28,b:1,c:ORNG,al:'c'});
  T(s,L.theFeatures,8.53,3.24,2.11,.3,{sz:12,c:SAGE,al:'c'});
  T(s,L.youWillCease,8.38,3.49,2.41,.91,{sz:11,c:INK,al:'c',lh:1.5});
  T(s,L.getTheDiscount,8.53,5.23,2.11,.3,{sz:12,i:1,c:SAGE,al:'c'});
  T(s,'PREMIUM',8.68,2.6,1.8,.51,{sz:24,b:1,c:ORNG,al:'c'});
  T(s,'$58,10',5.53,4.61,2.25,.57,{sz:28,b:1,c:W,al:'c'});
  T(s,L.theFeatures,5.6,3.14,2.11,.3,{sz:12,c:W,al:'c'});
  T(s,L.youWillCease,5.33,3.4,2.66,.91,{sz:11,c:W,al:'c',lh:1.5});
  T(s,L.getTheDiscount,5.6,5.19,2.11,.3,{sz:12,i:1,c:W,al:'c'});
  T(s,'MEDIUM',5.79,2.48,1.74,.51,{sz:24,b:1,c:W,al:'c'});
  button(s,2.83,5.7,'BOOKED / BUY');
  button(s,8.78,5.7,'BOOKED / BUY');
  button(s,5.81,5.7,'BOOKED / BUY',SAGE);
  brandTag(s,6.5);
  published(s,6.46);
}

function slide35(s) {
  draw(s, [
    ['blobBig',.21,3.74,5.16,3.76,SAND], ['blobCorner',0,0,2.48,1.78,CREAM2,{fh:1,fv:1,r:180}],
    ['fernA',1.11,5.04,.97,1.48,ORNG,{fh:1}], ['sparkle',1.9,2.38,.4,.67,ORNG,{fh:1}],
    ['sparkle',10.72,2.36,.4,.67,ORNG,{fh:1}], ['fernB',-.27,.38,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  rep(s,'ellipse',3.45,3.97,W,[[1.27,2.71],[5,2.71],[8.71,2.71]]);
  monstera(s,11.08,.92,2.06,1.71,CREAM2);
  rep(s,'roundRect',2.4,1.5,W,[[1.75,4.1],[5.48,4.1],[9.19,4.1]]);
  rep(s,'roundRect',2.24,1.27,SAND3,[[1.83,4.21],[5.56,4.21],[9.27,4.21]]);
  repT(s,2.25,.57,{sz:28,b:1,c:ORNG,al:'c'},[[[['$10/'],['day',{sz:20,b:0}]],1.82,3.55],[[['$20/'],
    ['day',{sz:20,b:0}]],5.55,3.55],[[['$30/'],['day',{sz:20,b:0}]],9.27,3.55]]);
  repT(s,1.6,.44,{sz:20,b:1,c:SAGE,al:'c'},[['BASIC',2.15,3.09],['Medium',5.88,3.09],['Premium',9.59,3.09]]);
  repT(s,2.2,1.19,{sz:11,c:SAGE2,al:'c',lh:1.5},[[L.youWillCease,1.85,4.25],[L.youWillCease,5.58,4.25],
    [L.youWillCease,9.29,4.25]]);
  button(s,2.1,5.75,'BOOKED / BUY',SAGE);
  button(s,5.83,5.75,'BOOKED / BUY',SAGE);
  button(s,9.54,5.75,'BOOKED / BUY',SAGE);
  T(s,L.chooseYourFavori,2.37,1.04,8.59,.71,{sz:36,b:1,c:ORNG,al:'c'});
  T(s,L.allAspectsOf,3.56,1.74,6.22,.78,{sz:14,c:ORNG,al:'c',lh:1.5,sa:6});
}

function slide36(s) {
  draw(s, [
    ['blobCorner',0,-.01,5.74,4.11,CREAM2], ['blobCorner',11.31,.34,2.35,1.68,CREAM2,{r:90}],
    ['roundRect',1.37,2,3.6,4.14,W,{rr:.248}], ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}],
    ['fernA',.92,4.95,.97,1.48,ORNG,{fh:1}],
  ]);
  monstera(s,3.94,1.25,2.06,1.71,CREAM2);
  T(s,'Strength Yoga Meditations',6.15,1.97,3.8,1.31,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf,6.15,3.51,5.31,.78,{sz:14,c:ORNG,lh:1.5,sa:6});
  T(s,L.theMoreYou,6.15,4.59,5.06,1.29,{sz:12,c:CREAM,lh:1.5,sa:6});
}

function slide37(s) {
  draw(s, [
    ['blobWave',9.36,0,3.98,2.81,CREAM2,{fh:1}], ['blobSoft',.28,4.97,2.24,2.81,SAND,{r:90}],
    ['fernB',11.51,1.11,1.1,1.52,ORNG2,{fh:1,r:42.8}], ['checkMark',1.19,2.72,.26,.26,SAND3],
    ['checkMark',1.19,4.38,.26,.26,SAND3],
  ]);
  monstera(s,1.63,6.21,1.08,.9,CREAM2);
  T(s,'Weakness Yoga Meditations',1.19,1.09,3.97,1.31,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf3,1.61,4.29,5.58,1.89,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.todayModernPhysi,1.61,2.56,5.58,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
}

function slide38(s) {
  draw(s, [
    ['leaf28',11.3,4.61,2.04,2.91,SAND,{fv:1}], ['blobCorner',-.54,4.22,3.82,2.74,CREAM2,{fh:1,fv:1,r:90}],
    ['blobSoft',11.55,-.2,1.58,1.98,SAND,{fh:1,fv:1,r:90}], ['ellipse',1.32,1.94,3.45,3.97,W],
    ['fernA',1.2,5.85,.97,1.48,ORNG,{fh:1}], ['fernB',4.12,2.23,1.1,1.52,ORNG2,{fh:1,r:42.8}],
    ['ellipse',5.48,2.73,.88,.88,SAND,{a:75,ln:SAND}], ['ellipse',5.57,2.83,.69,.69,SAND],
    ['ellipse',5.48,4.73,.88,.88,SAND,{a:75,ln:SAND}], ['ellipse',5.57,4.82,.69,.69,SAND],
  ]);
  T(s,'Opportunity Yoga Meditations',5.57,1.32,4.39,1.31,{sz:36,b:1,c:ORNG});
  T(s,'01',5.57,2.83,.69,.69,{sz:14,c:SAGE2,al:'c',va:1});
  T(s,'Peacefull Mind',6.5,2.73,1.77,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.allAspectsOf2,6.5,3.17,4.78,.98,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.theMoreYou,6.5,5.1,4.78,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,'02',5.57,4.82,.69,.69,{sz:14,c:SAGE2,al:'c',va:1});
  T(s,L.balanceBody,6.5,4.73,1.77,.37,{sz:16,b:1,c:ORNG2});
}

function slide39(s) {
  draw(s, [
    ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}], ['blobWave',0,0,1.74,1.23,SAND],
    ['blobCorner',9.75,0,3.58,2.56,CREAM2,{fh:1}], ['roundRect',.93,3.99,8.03,2.6,W,{rr:.131}],
    ['roundRect',.93,1.12,8.03,2.6,W,{rr:.131}], ['fernA',.59,5.92,.97,1.48,ORNG,{fh:1}],
    ['fernB',12.09,1.62,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  rep(s,'ellipse',.36,.36,ORNG,[[1.65,4.14],[1.65,4.85],[4.21,4.14],[4.21,4.85]]);
  rep(s,'chevron',.08,.14,W,[[1.8,4.25],[1.8,4.96],[4.36,4.25],[4.36,4.96]],{ln:W});
  T(s,'Threat Yoga Meditations',1.39,1.45,6.05,.71,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf,1.39,2.27,5.31,.78,{sz:14,c:ORNG,lh:1.5,sa:6});
  repT(s,1.77,.37,{sz:16,b:1,c:SAGE2},[[L.balanceBody,2.1,4.17],['Peacefull Mind',2.1,4.85],
    [L.healthyLife,4.66,4.17]]);
  T(s,'Drop Blood Preasure',4.66,4.85,2.48,.37,{sz:16,b:1,c:SAGE2});
  T(s,L.theMoreYou,1.61,5.48,6.45,.98,{sz:12,c:INK,lh:1.5,sa:6});
}

function slide40(s) {
  draw(s, [
    ['leaf28',11.62,-.02,1.71,2.45,SAND], ['blobCorner',-.54,.54,3.82,2.74,CREAM2,{fh:1,r:270}],
    ['blobSoft',11.79,5.96,1.37,1.71,SAND,{fh:1,r:270}], ['fernA',.59,5.92,.97,1.48,ORNG,{fh:1}],
    ['fernB',12.09,1.62,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  rep(s,'ellipse',2.22,2.22,W,[[4.18,1.79],[6.97,1.79],[9.73,1.79],[1.47,1.79]]);
  ring(s,2.54,2.92,.88,.194,CREAM);
  ring(s,5.29,2.92,.88,.194,SAND);
  ring(s,8.08,2.92,.88,.194,CREAM);
  ring(s,10.83,2.92,.88,.194,SAND);
  T(s,L.chartsPlans,3.91,.77,5.04,.71,{sz:36,b:1,c:ORNG,al:'c'});
  repT(s,1.57,.3,{sz:12,c:ORNG,al:'c'},[['Data 01',1.75,3.05],['Data 02',4.5,3.05],['Data 01',7.3,3.05],
    ['Data 02',10.05,3.05]]);
  repT(s,1.57,.64,{sz:32,b:1,c:ORNG,al:'c'},[['86%',1.75,2.53],['65%',4.5,2.53],['96%',7.3,2.53],
    ['35%',10.05,2.53]]);
  repT(s,1.58,.38,{sz:14,c:CREAM,al:'c',lh:1.2},[[L.expertService,1.83,4.59],[L.expertService,7.33,4.59],
    [L.expertService,4.46,4.59],[L.expertService,10.12,4.59]]);
  repT(s,2.09,.46,{sz:18,b:1,c:ORNG,al:'c',lh:1.2},[['Charts Yoga',1.57,4.18],['Charts Yoga',7.08,4.18],
    ['Charts Yoga',4.21,4.18],['Charts Yoga',9.87,4.18]]);
  repT(s,2.54,.98,{sz:12,c:CREAM,al:'c',lh:1.5,sa:6},[[L.youWillCease,1.35,5.07],[L.youWillCease,6.86,5.07],
    [L.youWillCease,3.98,5.07],[L.youWillCease,9.64,5.07]]);
}

function slide41(s) {
  draw(s, [
    ['blobCorner',9.06,.71,4.98,3.56,CREAM2,{r:90}], ['blobCorner',0,.01,2.07,1.48,CREAM2],
    ['sparkle',1.04,.94,.4,.67,ORNG,{fh:1}], ['sparkle',11.26,2.31,.4,.67,ORNG,{fh:1}],
    ['roundRect',6.95,3.44,.43,2.02,SAND,{r:180,rr:.215}],
    ['roundRect',8.07,3.6,.43,1.86,SAGE,{r:180,rr:.215}], ['roundRect',9.2,2.7,.43,2.76,SAND,{r:180,rr:.215}],
    ['roundRect',10.32,2.9,.43,2.56,SAGE,{r:180,rr:.215}],
    ['roundRect',11.44,4.6,.43,.86,SAND,{r:180,rr:.215}],
  ]);
  monstera(s,.19,5.21,2.06,1.71,CREAM2);
  rep(s,'roundRect',.43,2.95,'ADB9CA',[[6.95,2.52],[8.07,2.52],[9.2,2.52],[10.32,2.52],[11.44,2.52]],{a:76,r:180,rr:.215});
  repT(s,.69,.34,{sz:12,c:CREAM,al:'c',lh:1.2},[['65%',6.82,2.04],['63%',7.94,2.04],['97%',9.06,2.04],
    ['91%',10.19,2.04],['24%',11.31,2.04]]);
  T(s,L.chartsPlans,.97,1.69,3.28,.71,{sz:36,b:1,c:ORNG});
  T(s,L.allAspectsOf,.97,2.52,5.31,.78,{sz:14,c:ORNG,al:'j',lh:1.5,sa:6});
  T(s,L.theMoreYou,.97,3.6,5.31,1.29,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6});
  button(s,.97,5.22,'Discover More');
}

function slide42(s) {
  draw(s, [
    ['blobWave',0,0,2.34,1.65,CREAM2], ['blobSoft',10.23,4.4,2.75,3.44,SAND,{fh:1,r:270}],
    ['fernA',.52,.7,.97,1.48,ORNG,{fh:1}], ['fernB',11.98,4.54,1.1,1.52,ORNG2,{fh:1,r:42.8}],
    ['leaf29',9.36,2.69,2.35,.78,SAGE], ['leaf30',6.39,3.24,1.14,.76,SAGE], ['leaf31',5.4,2.13,1.05,.91,SAGE],
    ['bit20',7.24,4.46,.58,.95,SAGE], ['bit21',7.38,4.87,.25,1.08,SAGE], ['bit22',7.35,5.01,.21,1.04,SAGE],
    ['bit23',7.51,1.78,.59,.46,SAGE], ['leaf32',6.2,2.16,1.6,1.4,SAGE], ['bit24',6.39,3.8,.63,.47,SAGE],
    ['bit25',7.4,5.1,.35,.83,SAGE], ['bit26',10.66,3.22,.68,.32,SAGE], ['bit27',8.97,2.62,.44,.52,SAGE],
    ['bit28',11.33,4.86,.88,.79,SAGE], ['leaf33',10.29,3.11,1.35,1.02,SAGE], ['bit29',9.82,3.16,.83,.44,SAGE],
    ['bit30',10.37,3.81,.62,.61,SAGE], ['bit31',9.8,3.66,.45,.36,SAGE], ['bit32',9.63,3.85,.46,.39,SAGE],
    ['bit33',9.36,4.1,.36,.42,SAGE], ['bit34',8.7,3.74,.44,.43,SAGE], ['bit35',9.15,4.48,.42,.42,SAGE],
    ['leaf34',9.38,2.05,2.89,1.56,CREAM], ['leaf35',7.8,1.84,.93,1.19,CREAM],
    ['bit36',10.98,4.44,1.22,.41,CREAM], ['leaf36',6.28,3.28,1.24,.74,CREAM],
    ['bit37',7.28,4.48,.86,.91,CREAM],
  ]);
  T(s,L.chartsPlans,7.61,.94,3.28,.71,{sz:36,b:1,c:ORNG,al:'c'});
  repT(s,1.39,.71,{sz:36,b:1,c:CREAM,al:'r'},[['01.',.56,1.64],['02.',.56,3.34],['03.',.52,5.27]]);
  repT(s,1.58,.38,{sz:14,c:CREAM,lh:1.2},[[L.expertService,1.95,2.02],[L.expertService,1.95,3.72],
    [L.expertService,1.91,5.65]]);
  T(s,L.balanceBody,1.95,1.6,2.09,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.meditationTime,1.95,3.3,2.32,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  T(s,L.peacefulMinds,1.91,5.24,2.09,.46,{sz:18,b:1,c:ORNG,lh:1.2});
  repT(s,3.6,.68,{sz:12,c:CREAM,al:'j',lh:1.5,sa:6},[[L.youWillCease,2.03,2.48],[L.youWillCease,2.03,4.2],
    [L.youWillCease,2.03,6.13]]);
}

function slide43(s) {
  draw(s, [
    ['blobCorner',9.06,.71,4.98,3.56,CREAM2,{r:90}], ['blob12',-.03,4.64,2.17,2.86,SAND,{fh:1,fv:1,r:180}],
    ['blobSoft',11.55,5.72,1.58,1.98,SAND,{fh:1,r:270}], ['fernB',11.04,.77,1.25,1.74,ORNG2,{fh:1,r:338.6}],
    ['fernA',11.67,5.45,1.35,2.05,ORNG,{fh:1,fv:1,r:180}], ['leaf37',7.5,3.35,2.09,2.09,SAND],
    ['leaf38',8.61,2.24,2.09,1.82,SAGE], ['leaf39',9.72,3.35,2.09,2.09,ORNG],
    ['checkMark',1.3,3.58,.26,.26,ORNG], ['checkMark',1.3,4.46,.26,.26,ORNG],
  ]);
  monstera(s,.87,5.06,2.06,1.71,CREAM2);
  T(s,'Peacefull Mind',1.3,2.99,1.77,.37,{sz:16,b:1,c:ORNG2});
  T(s,L.youWillCease,1.72,3.42,4.4,.78,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.todayModernPhysi4,1.72,4.3,5.25,1.13,{sz:14,c:CREAM,al:'j',lh:1.5,sa:6});
  T(s,L.chartsPlans,1.3,2.06,3.28,.71,{sz:36,b:1,c:ORNG});
}

function slide44(s) {
  draw(s, [
    ['blobBig',.1,-0,6.3,4.6,SAND,{fv:1}], ['blobCorner',11.26,.01,2.07,1.48,CREAM2,{fh:1}],
    ['roundRect',1.75,1.93,3.61,3.77,W,{rr:.182}], ['blobCorner',-.54,4.22,3.82,2.74,CREAM2,{fh:1,fv:1,r:90}],
    ['fernA',1.09,4.83,.97,1.48,ORNG,{fh:1}], ['fernB',4.86,4.78,1.1,1.52,ORNG2,{fh:1,r:42.8}],
  ]);
  monstera(s,4.99,.21,2.06,1.71,CREAM2);
  social(s,11.55,6.91);
  T(s,[['Adreess'],['	 : 76 United Street Ca,cCalifornia, CA 62 71593',{b:0,br:1}],['Phone'],
    ['	 : +09 235 854 787  ( +65 235 854 787 )',{b:0,br:1}],['Mail'],
    [' 	 : @yogaes.office.official.com',{b:0,br:1}],['Website'],
    ['	 : @yogaes.office.com',{b:0}]],6.67,3.94,5.62,1.49,{sz:14,b:1,c:CREAM,lh:1.5});
  T(s,'Contact Us For More Information,',6.67,2.44,4.87,1.31,{sz:36,b:1,c:ORNG});
  T(s,'MORE INFORMATION',2.41,4.29,2.21,.3,{sz:12,c:ORNG,al:'c'});
  T(s,'We 24 Hours Ready For Any Call Of Yoga',1.92,4.65,3.19,.77,{sz:20,b:1,c:SAGE,al:'c'});
  brandTag(s,6.5);
  published(s,6.46);
}

function slide45(s) {
  draw(s, [
    ['roundRect',.98,2.44,11.38,3.16,W], ['roundRect',1.23,2.68,10.88,2.69,SAND2],
    ['blobCorner',-0,0,2.31,1.65,CREAM2], ['blobCorner',11.15,5.93,2.19,1.56,CREAM2,{r:180}],
    ['blobBig',1.18,4.24,4.1,3.27,SAND,{fh:1}], ['fernB',3.48,3.92,1.64,2.28,ORNG2,{r:21.4}],
    ['fernB',1.61,4.46,1.18,1.64,ORNG2,{r:317.2}], ['sparkle',.49,4.19,.37,.63,ORNG],
    ['sparkle',3.88,1.69,.37,.63,ORNG], ['bit1',2.77,6.73,1.04,.32,'291701'],
    ['leaf1',2.75,6.84,1.4,.61,SAND4], ['leaf2',1.95,6.54,1.05,.7,BROWN], ['leaf3',2.43,6.84,1.4,.61,SAND4],
    ['leaf4',3.58,6.54,1.05,.7,BROWN], ['leaf5',2.77,4.83,1.04,1.89,ORNG], ['leaf6',3.3,4.85,1.3,1.17,SAND4],
    ['leaf7',1.99,4.85,1.29,1.17,SAND4], ['leaf8',2.96,3.71,.65,1,SAND4], ['bit2',2.95,3.71,.68,.48,BROWN],
    ['fernA',.41,5.58,1.26,1.92,ORNG,{fv:1,r:180}], ['fernA',4.66,6.12,.91,1.38,ORNG],
    ['leafSpray',1.1,.93,1.75,1.64,SAND,{fh:1,r:270}], ['leafSpray',11.13,5.31,1.75,1.64,SAND,{fh:1,r:175.9}],
  ]);
  monstera(s,3.68,4.92,2.68,2.22,CREAM2);
  monstera(s,.59,5.23,2.16,1.79,CREAM2);
  social(s,11.55,.56);
  T(s,'THANKS',3.97,2.79,7.93,2.42,{sz:138,b:1,c:GOLD,al:'c'});
  brandTag(s,.15);
  T(s,L.healthyYogaFound,5.9,4.59,3.93,.44,{sz:20,c:ORNG,al:'c'});
  published(s,.11);
}

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40, slide41, slide42, slide43, slide44, slide45];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };

for (const build of SLIDES) {
  const s = pptx.addSlide();
  s.background = { color: BG };
  build(s);
}

pptx.writeFile({ fileName: path.join(__dirname, '08165b97-eaf6-438e-ae86-bd4ad9d03daa_grok_final.pptx') })
  .then(f => console.log('wrote', f));
