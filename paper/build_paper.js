/* Builds the scientific paper (.docx) from the result figures.  Run:  node build_paper.js
   Requires the `docx` npm package. Figures are read from ../results/figures. */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, ImageRun, PageBreak, Footer, PageNumber,
} = require("docx");

const FIG = path.join(__dirname, "..", "results", "figures");
const OUT = path.join(__dirname, "AI_Assisted_Web_Pentest_CICIDS2017.docx");
const MAX_W = 624; // px at 96 dpi = 6.5 in content width

// ---------------------------------------------------------------- helpers
const FONT = "Calibri";
const R = (text, o = {}) => new TextRun({ text, size: 21, bold: o.b, italics: o.i, font: FONT });
const p = (runs, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED, spacing: { after: o.after ?? 120, line: 276 },
  children: typeof runs === "string" ? [R(runs)] : runs,
});
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t })] });
const bullet = (runs) => new Paragraph({ numbering: { reference: "bullets", level: 0 },
  spacing: { after: 60, line: 264 }, children: typeof runs === "string" ? [R(runs)] : runs });
const cap = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 200 },
  children: [new TextRun({ text: t, italics: true, size: 18, font: FONT })] });

function figure(file, caption, widthFrac = 1.0) {
  const buf = fs.readFileSync(path.join(FIG, file));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const tw = Math.round(MAX_W * widthFrac), th = Math.round(tw * h / w);
  return [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 }, keepNext: true,
    children: [new ImageRun({ type: "png", data: buf, transformation: { width: tw, height: th } })] }),
    cap(caption)];
}

const HEAD = "1F3B57", ALT = "F1F5F9", HIGHLIGHT = "FFF4E5";
function cell(text, w, { head = false, bold = false, fill, align } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 80, right: 80 },
    shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined,
    children: [new Paragraph({ alignment: align || (head ? AlignmentType.CENTER : AlignmentType.LEFT),
      keepNext: true, keepLines: true,   // keep the whole table (and its caption) on one page
      children: [new TextRun({ text: String(text), bold: head || bold, size: 17, font: FONT,
        color: head ? "FFFFFF" : "000000" })] })],
  });
}
function table(widths, header, rows, { highlight = () => false } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const line = { style: BorderStyle.SINGLE, size: 4, color: "B8C2CC" };
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    borders: { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line },
    rows: [
      new TableRow({ tableHeader: true, children: header.map((t, i) => cell(t, widths[i], { head: true, fill: HEAD })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((t, i) => cell(t, widths[i], {
        bold: i === 0, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
        fill: highlight(ri) ? HIGHLIGHT : (ri % 2 ? ALT : undefined) })) })),
    ],
  });
}

// ---------------------------------------------------------------- content
const c = [];

// Title block
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
  children: [new TextRun({ text: "AI-Assisted Web Application Penetration Testing", bold: true, size: 34, font: FONT })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
  children: [new TextRun({ text: "Detecting and Classifying Web Attacks in CIC-IDS2017 Network Flows, and How AI Compares with Conventional Detection",
    bold: true, size: 25, font: FONT, color: "2E5A80" })] }));
c.push(p([R("[Your Name]  •  Student ID: [________]")], { align: AlignmentType.CENTER, after: 20 }));
c.push(p([R("CSCI561 — Computer and Network Security  •  November 2026", { i: true })], { align: AlignmentType.CENTER, after: 240 }));

c.push(new Paragraph({ spacing: { after: 80 }, border: { top: { style: BorderStyle.SINGLE, size: 6, color: HEAD, space: 6 } },
  children: [new TextRun({ text: "Abstract", bold: true, size: 22, font: FONT })] }));
c.push(p("Automated web-vulnerability scanners and intrusion detection systems traditionally rely on hand-written signatures and thresholds. This project investigates how machine learning (ML) can assist in detecting and classifying web-application attacks, using the web-attack day of the public CIC-IDS2017 dataset, and compares it with conventional detection. A reproducible pipeline downloads and verifies the data, removes non-finite and duplicate records, and restricts the analysis to HTTP flows, leaving 19,725 unique flows (17,582 benign, 1,470 brute-force, 652 cross-site scripting, 21 SQL injection) described by 66 CICFlowMeter statistics. Five models — logistic regression, a linear SVM, random forest, gradient-boosted trees and a multi-layer perceptron — are compared with a conventional statistical-threshold IDS on a held-out test set and with 5-fold cross-validation. ML detects web attacks almost perfectly (F1 up to 0.998, ROC-AUC 1.000), whereas the threshold IDS detects none at a 1% false-alarm budget: the attack flows are statistically ordinary. Naming the attack type is much harder (macro-F1 at most 0.78). A validity audit explains why: 91% of the flows labelled as attacks carry no payload at all, and 94% share the attacker’s TCP initial window, a host fingerprint. With both confounds removed, the best detector still reaches F1 0.968 and attack-type macro-F1 rises to 0.90. Robustness tests expose the limits: an attack type absent from training (SQL injection) is missed far more often, slowing an attack down two- to ten-fold defeats the tree ensembles, and on an unseen benign day the false-alarm rate of the linear models more than doubles. AI substantially outperforms threshold rules on flow data, but headline accuracy on this dataset overstates real-world performance; flow-based AI complements, rather than replaces, payload-based scanning such as OWASP ZAP, for which a lab harness is provided."));
c.push(p([R("Keywords: ", { b: true }), R("web application security, intrusion detection, machine learning, CIC-IDS2017, network flows, shortcut learning, OWASP ZAP", { i: true })]));
c.push(new Paragraph({ children: [new PageBreak()] }));

// 1 Introduction
c.push(h1("1. Introduction"));
c.push(p("Web applications are exposed to a constant stream of attacks such as credential brute-forcing, cross-site scripting (XSS) and SQL injection (SQLi). Defenders traditionally rely on two kinds of conventional tooling: automated vulnerability scanners and web application firewalls, which match payloads against hand-written signatures, and network intrusion detection systems (IDS), which raise alarms when traffic statistics leave a configured normal range. Both are transparent and need no training data, but signatures miss variants and must be maintained by hand, and fixed thresholds struggle when attacks look like normal traffic."));
c.push(p("Machine learning promises to learn the boundary between benign and malicious traffic directly from labelled examples. Public flow datasets such as CIC-IDS2017 [1] have made this easy to try, and near-perfect accuracies are routinely reported. At the same time, audits of CIC-IDS2017 have found construction and labelling errors [2], [3], and the security ML literature warns that spurious correlations can make a model look excellent for the wrong reasons [6], [7]. This project therefore asks:"));
c.push(p([R("How effectively can AI/ML assist in detecting and classifying web-application attacks, and how does its performance compare with conventional detection?", { i: true })]));
c.push(p("The contributions are:"));
c.push(bullet("A reproducible data pipeline: a verified downloader (official source with a pinned mirror fallback), documented cleaning, and an HTTP-only scope that removes the trivial port-number shortcut."));
c.push(bullet("A comparison of five ML models with a conventional threshold IDS for binary detection and attack-type classification, on held-out data and with 5-fold cross-validation."));
c.push(bullet("A validity audit that quantifies payload-less artifact flows and a TCP-fingerprint shortcut, and re-evaluates every model with both controlled."));
c.push(bullet("Robustness tests: attack types never seen in training, low-and-slow timing evasion, false alarms on an unseen benign day, and the effect of the evaluation scope."));
c.push(bullet("An AI-assisted assessment prototype with an OWASP ZAP lab harness (packet capture → CICFlowMeter → scoring → comparison report), validated end to end on flow exports from both CICFlowMeter implementations."));

// 2 Background
c.push(h1("2. Background and Related Work"));
c.push(h2("2.1 Web attacks and flow records"));
c.push(p("CIC-IDS2017 [1] was captured over five working days in a testbed with realistic background traffic. On Thursday morning the attackers ran three automated web attacks against an intentionally vulnerable web application: brute-forcing a login form, injecting XSS payloads and performing SQL injection. The released CSV files do not contain packets or payloads. Each row is a bidirectional flow summarised by CICFlowMeter into about 80 statistics: packet and byte counts per direction, packet-length statistics, durations and inter-arrival times (IAT), rates, TCP flag counts, header lengths and initial TCP window sizes."));
c.push(h2("2.2 Conventional detection"));
c.push(p("Conventional web scanners such as OWASP ZAP and Burp Suite actively send attack payloads to an application and confirm vulnerabilities from its responses, and WAF rule sets match known payload patterns in requests. Neither can operate on flow records, which contain no payload. The conventional technique that does apply to flow statistics is statistical anomaly detection, introduced by Denning [8]: a profile of normal behaviour is built for each metric and an alarm is raised when an observation falls outside it. This project uses such a detector as its conventional baseline (Section 4.3) and covers active scanning with a separate lab harness (Section 5.10)."));
c.push(h2("2.3 Machine learning for intrusion and web-attack detection"));
c.push(p([R("Sharafaldin "), R("et al.", { i: true }), R(" [1] introduced CIC-IDS2017 and characterised its attack traffic. Maseer "), R("et al.", { i: true }),
  R(" [4] benchmarked a range of supervised and unsupervised algorithms on it for anomaly-based intrusion detection, and Kurniabudi "), R("et al.", { i: true }),
  R(" [5] analysed its features with information gain to select relevant subsets. Data quality has since been questioned. Engelen "), R("et al.", { i: true }),
  R(" [2] showed that CICFlowMeter splits connections at the first FIN into spurious “TCP appendix” flows, and that most brute-force and XSS flows carry no payload (1,507 brute-force flows shrink to 151, and 652 XSS flows to 27, after a payload filter). Lanvin "), R("et al.", { i: true }),
  R(" [3] catalogued further errors and showed that correcting them significantly changes measured detection performance. More broadly, Arp "), R("et al.", { i: true }),
  R(" [6] describe pitfalls such as spurious correlations and inappropriate baselines in security ML, and Sommer and Paxson [7] explain why network intrusion detection is harder for ML than other domains. For web attacks specifically, many studies learn from HTTP content rather than flows, for example deep models over request strings [9] or classifiers over XSS payload features [10]. This project tests what can be achieved from flow statistics alone and audits the result.")]));

// 3 Data
c.push(h1("3. Dataset and Preprocessing"));
c.push(h2("3.1 Acquisition and integrity"));
c.push(p("The project uses the “MachineLearningCVE” CSV release of CIC-IDS2017: the Thursday-morning file, which contains the web attacks, and the Monday file, which contains only benign traffic. The script download_data.py first tries the official archive from the Canadian Institute for Cybersecurity. If that fails, it falls back to a public GitHub mirror cloned with a blob-less partial clone pinned to a fixed commit, so only the required archive parts (≈136 MB) are transferred. Both files are then verified against the published label counts: Thursday has 170,366 flows (168,186 benign, 1,507 brute force, 652 XSS, 21 SQL injection) and Monday has 529,918 benign flows. Their SHA-256 digests are also recorded."));
c.push(h2("3.2 Cleaning"));
c.push(p("The labels were repaired (a mis-encoded dash appears as “Web Attack � Brute Force”), whitespace was stripped from column names, and the column “Fwd Header Length.1” was removed because it exactly duplicates “Fwd Header Length”. 135 zero-duration flows have undefined rates (infinite Flow Bytes/s and Flow Packets/s). All of them are benign, and they were dropped. 6,052 exact duplicate rows were removed so that identical flows cannot appear in both the training and test sets, and 10 constant columns were removed. This leaves 164,179 unique flows. No identical feature vector carries two different labels."));
c.push(h2("3.3 Scope: HTTP flows"));
c.push(p("All 2,143 unique attack flows target the web server on port 80, while 89% of benign flows use other ports. A detector trained on all traffic can therefore succeed largely by recognising web traffic. Section 5.9 shows that a rule using only the destination port already reaches ROC-AUC 0.947. To evaluate whether models can tell malicious from benign web traffic, the analysis keeps only HTTP flows and removes the then-constant port feature. The final dataset has 19,725 flows (10.9% attacks) and 66 features (Table 1, Fig. 1). Monday was processed identically, giving 47,799 unique benign HTTP flows for the unseen-day test."));
c.push(table([2700, 1700, 1700, 1600, 1600], ["Class", "Thursday raw", "HTTP scope (unique)", "Train (80%)", "Test (20%)"], [
  ["BENIGN", "168,186", "17,582", "14,066", "3,516"],
  ["Brute Force", "1,507", "1,470", "1,176", "294"],
  ["XSS", "652", "652", "521", "131"],
  ["SQL Injection", "21", "21", "17", "4"],
  ["Total", "170,366", "19,725", "15,780", "3,945"],
]));
c.push(cap("Table 1. Class counts before and after cleaning, and the held-out split (stratified by attack type)."));
c.push(...figure("fig01_data.png", "Fig. 1. Data funnel for Thursday-morning traffic (left) and the class distribution of the final HTTP-flow dataset (right, log scale)."));

// 4 Methodology
c.push(h1("4. Methodology"));
c.push(h2("4.1 Features"));
c.push(p("Each flow is described by 66 CICFlowMeter statistics (Appendix B). They cover volume (packets and bytes per direction, sub-flows), packet-length statistics, timing (duration, forward, backward and flow IAT, active and idle periods), rates, TCP flag counts, header lengths and TCP initial window sizes. The statistics are heavy-tailed, so the linear models and the MLP receive a signed log transform, sign(x)·log(1+|x|), followed by standardisation. The transform keeps the −1 sentinel used for unobserved windows. Tree models are scale-invariant and use raw values."));
c.push(h2("4.2 AI/ML models"));
c.push(p("Five supervised models span the suggested families: logistic regression (C = 1) and a linear SVM (C = 0.5); a random forest with 300 trees; gradient-boosted trees (scikit-learn’s HistGradientBoosting, 300 iterations, learning rate 0.1, early stopping); and a multi-layer perceptron with two hidden layers (128 and 64 units, early stopping). Attacks are rare, so every model is class-balanced, using class weights or, for the MLP, balanced sample weights. XGBoost could not be installed in the build environment, so HistGradientBoosting represents the same gradient-boosting family, and the code accepts XGBoost as a drop-in replacement. The same architectures are used for the binary task (benign vs. attack) and the four-class task (benign, brute force, XSS, SQLi)."));
c.push(h2("4.3 Conventional baseline: a threshold IDS"));
c.push(p("Following Denning’s statistical model [8], the baseline learns the normal range of every feature from benign training flows only, as the interval between the 0.5th and 99.5th percentiles. A flow’s anomaly score is the number of features that fall outside their range. The alarm threshold is the smallest score whose false-alarm rate on benign training flows is at most 1%, mimicking how an analyst tunes rules to an alert budget. This gives an alarm at eight or more out-of-range metrics (0.78% false alarms on training data). The detector never learns from attacks, and every alarm is explainable by the list of metrics that fired. It raises alarms but cannot name the attack type."));
c.push(h2("4.4 Evaluation protocol"));
c.push(p("A single stratified 80/20 split, stratified by attack type so that all four classes appear in both parts, provides the held-out test set (15,780 training and 3,945 test flows) for tables, ROC curves and confusion matrices. Stratified 5-fold cross-validation provides means and standard deviations. Its pooled out-of-fold predictions test every flow exactly once, including all 21 SQLi flows, which a single split cannot do (only 4 are in the test set). Metrics are accuracy, precision, recall (detection rate), F1, ROC-AUC, PR-AUC, the Matthews correlation coefficient, false-positive and false-negative rates, and the detection rate at fixed false-alarm budgets of 1% and 0.1%. Multi-class results use macro-F1, which weights every attack type equally."));
c.push(h2("4.5 Robustness and validity experiments"));
c.push(bullet([R("E3, unseen attack types: ", { b: true }), R("each attack type is removed from training in turn; detection of all its flows is compared with the out-of-fold detection rate when it was seen.")]));
c.push(bullet([R("E4, timing evasion: ", { b: true }), R("held-out attack flows are slowed down in feature space (time features ×k, rate features ÷k, k = 1 to 100) and re-scored without retraining.")]));
c.push(bullet([R("E5, unseen day: ", { b: true }), R("Thursday-trained detectors are run on all 47,799 benign Monday HTTP flows. Every alarm is false.")]));
c.push(bullet([R("E6, evaluation scope: ", { b: true }), R("results on all traffic, with and without the port feature, are compared with the HTTP-only scope.")]));
c.push(bullet([R("E7, validity audit: ", { b: true }), R("models are re-evaluated (5-fold CV) with TCP-fingerprint features removed, with payload-less flows removed, and with both.")]));

// 5 Results
c.push(h1("5. Results"));
c.push(h2("5.1 Detecting web attacks"));
c.push(p("All five models detect web attacks with F1 of at least 0.95 on the held-out test set (Table 2). The tree ensembles are near-perfect: gradient-boosted trees reach F1 0.997 with a false-positive rate of 0.03%, and the random forest raises no false alarm at all. Cross-validation agrees (gradient-boosted trees 0.998 ± 0.002, random forest 0.993 ± 0.002; Fig. 2, left). The ROC curves on a logarithmic false-positive axis (Fig. 3) reveal a practical difference. Under a strict 0.1% alarm budget the tree ensembles still detect 99.5–100% of attacks and the MLP 97%, but the linear models detect only 4–12%. They need about a 1% false-alarm rate to work."));
c.push(table([2250, 680, 680, 680, 720, 760, 760, 820, 820], ["Detector", "Prec.", "Recall", "F1", "ROC-AUC", "FPR", "TPR @1%", "TPR @0.1%", "CV F1"], [
  ["Logistic Regression", "0.918", "0.988", "0.952", "0.998", "1.08%", "0.988", "0.044", "0.959 ± 0.006"],
  ["Linear SVM", "0.940", "0.988", "0.964", "0.998", "0.77%", "0.991", "0.117", "0.973 ± 0.005"],
  ["Random Forest", "1.000", "0.986", "0.993", "1.000", "0.00%", "1.000", "0.995", "0.993 ± 0.002"],
  ["Gradient-Boosted Trees", "0.998", "0.995", "0.997", "1.000", "0.03%", "1.000", "1.000", "0.998 ± 0.002"],
  ["Neural Network (MLP)", "0.966", "0.991", "0.978", "1.000", "0.43%", "1.000", "0.972", "0.986 ± 0.010"],
  ["Threshold IDS (conv.)", "0.000", "0.000", "0.000", "0.490", "0.68%", "0.000", "0.000", "0.000 ± 0.000"],
], { highlight: (i) => i === 5 }));
c.push(cap("Table 2. Binary detection on the held-out test set (3,945 HTTP flows, 429 attacks); the last column is the 5-fold CV mean ± std. TPR @x% = detection rate at an x% false-alarm budget."));
c.push(...figure("fig02_detect_vs_classify.png", "Fig. 2. Detecting that an attack is happening (left) is far easier than naming which attack it is (right). 5-fold CV mean ± std."));
c.push(...figure("fig03_roc.png", "Fig. 3. ROC curves on a log-scaled false-positive axis. Tree ensembles detect attacks even at very low alarm budgets; linear models need about 1% false alarms.", 0.72));

c.push(h2("5.2 Why the conventional threshold IDS fails"));
c.push(p("The threshold IDS detects none of the 429 test attacks at its 1% budget, and its ROC-AUC of 0.49 is no better than chance (Fig. 4). Attack flows are not statistical outliers. On every metric they sit near the middle of the benign distribution: for example, the median attack flow is at the 49th benign percentile for flow duration and the 52nd for forward packet count. No attack flow has more than five out-of-range metrics, while staying within 1% false alarms requires eight. The benign tail of large transfers and long sessions is far more “anomalous” than the attacks. Web attacks in this data are therefore ordinary on each metric taken separately and separable only by combinations of metrics, which is exactly where learned models add value over per-metric rules."));
c.push(...figure("fig04_threshold_ids_scores.png", "Fig. 4. Distribution of the threshold IDS anomaly score. Attack flows never reach the alarm threshold that a 1% false-alarm budget requires.", 0.85));

c.push(h2("5.3 Classifying the attack type"));
c.push(p("Naming the attack is much harder than detecting it (Fig. 2, right; Table 3). The best model, gradient-boosted trees, reaches a macro-F1 of 0.78 ± 0.04, and the others range from 0.52 to 0.68. The out-of-fold confusion matrix (Fig. 6, left) shows the problem: 35% of brute-force flows are labelled XSS and 35% of XSS flows are labelled brute force. SQL injection, despite having only 21 flows, is recognised better (F1 0.84 for gradient-boosted trees). Section 5.4 explains why brute force and XSS are confused."));
c.push(table([2350, 1250, 1100, 1250, 1000, 1250], ["Model", "Macro-F1 (CV)", "BENIGN", "Brute Force", "XSS", "SQL Injection"], [
  ["Logistic Regression", "0.572 ± 0.008", "0.991", "0.530", "0.488", "0.280"],
  ["Linear SVM", "0.539 ± 0.013", "0.997", "0.797", "0.065", "0.296"],
  ["Random Forest", "0.680 ± 0.054", "0.999", "0.776", "0.381", "0.562"],
  ["Gradient-Boosted Trees", "0.777 ± 0.039", "0.999", "0.720", "0.532", "0.844"],
  ["Neural Network (MLP)", "0.520 ± 0.058", "0.985", "0.704", "0.230", "0.164"],
], { highlight: (i) => i === 3 }));
c.push(cap("Table 3. Attack-type classification: 5-fold CV macro-F1 and per-class F1 from pooled out-of-fold predictions (all 19,725 flows)."));

c.push(h2("5.4 Validity audit: what are the models actually detecting?"));
c.push(p("The audit (E7) measured two confounds. First, 89.7% of brute-force, 96.3% of XSS and 42.9% of SQLi flows carry no payload in either direction, compared with 15.4% of benign flows. Only 187 of the 2,143 attack flows (8.7%) contain a request from the attacker. Most of the others are 3-packet connection fragments (1,818 of 1,956), the TCP-appendix artifacts described by Engelen et al. [2]. Our counts agree with theirs: 151 brute-force and 12 SQLi flows survive a payload filter in both studies (24 vs. 27 XSS flows, after our de-duplication). Second, 92.8% of brute-force and 96.5% of XSS flows share the TCP initial window 29,200 of the attacker’s Linux machine, compared with 25.3% of benign flows. Init_Win_bytes_forward (with the server-side window and minimum segment size) identifies the attacking host rather than the attack."));
c.push(p("Table 4 and Fig. 5 re-evaluate every model under four settings. Removing the fingerprint features costs little (S1). Removing payload-less flows (S2, S3) lowers the attack base rate to 1.2% and reduces F1, sharply for the linear models, whose precision falls to 0.59–0.70. The random forest still reaches F1 0.968 and gradient-boosted trees 0.947 in the most conservative setting (S3), so ML does detect real attack requests from flow statistics alone. Attack-type classification improves markedly once the artifacts are gone: macro-F1 rises from 0.77 to 0.90 for gradient-boosted trees and from 0.54 to 0.83 for the linear SVM (Fig. 6, right: brute force 97%, XSS 83%, SQLi 75% recall). The poor result in Section 5.3 was largely an artifact. An empty connection cannot reveal whether it belonged to a brute-force or an XSS run. The threshold IDS remains at zero detections in every setting."));
c.push(table([2150, 830, 830, 830, 830, 900, 900, 900, 900], ["Model", "F1 S0", "F1 S1", "F1 S2", "F1 S3", "Type S0", "Type S1", "Type S2", "Type S3"], [
  ["Logistic Regression", "0.959", "0.950", "0.770", "0.741", "0.572", "0.533", "0.717", "0.678"],
  ["Linear SVM", "0.973", "0.968", "0.823", "0.810", "0.539", "0.574", "0.828", "0.827"],
  ["Random Forest", "0.993", "0.988", "0.981", "0.968", "0.680", "0.631", "0.890", "0.875"],
  ["Gradient-Boosted Trees", "0.998", "0.990", "0.945", "0.947", "0.774", "0.655", "0.878", "0.896"],
  ["Neural Network (MLP)", "0.986", "0.974", "0.921", "0.909", "0.521", "0.516", "0.730", "0.737"],
]));
c.push(cap("Table 4. Validity audit (5-fold CV, pooled out-of-fold). F1 = binary detection; Type = attack-type macro-F1. S0 all HTTP flows; S1 without TCP-fingerprint features; S2 payload-bearing flows only (187 attacks); S3 both."));
c.push(...figure("fig05_validity_audit.png", "Fig. 5. Detection F1 (left) and attack-type macro-F1 (right) as confounds are removed, from S0 (standard) to S3 (most conservative)."));
c.push(...figure("fig06_confusion.png", "Fig. 6. Out-of-fold confusion matrices of gradient-boosted trees. The brute-force/XSS confusion of the standard setup (left) largely disappears once payload-less artifact flows are removed (right)."));

c.push(h2("5.5 Which features matter"));
c.push(p("The random forest’s impurity importance (Fig. 10, left) is spread over packet-size features (maximum packet length, average packet size, forward bytes) and Fwd IAT Min. The three TCP-fingerprint features rank 6th, 11th and 12th. Permutation importance on the test set (Fig. 10, right) tells a sharper story. Only Fwd IAT Min matters individually, because the other features are redundant with one another. A model that depends on a single timing statistic is fragile, which E4 confirms."));
c.push(...figure("fig10_feature_importance.png", "Fig. 10. Random-forest feature importance. Impurity importance (left; fingerprint features in orange) and permutation importance on the test set (right)."));

c.push(h2("5.6 Attack types never seen in training"));
c.push(p("When brute force or XSS is removed from training, the models still detect 89–97% of its flows (Fig. 7). This is not real zero-day generalisation, because the artifact flows of the two attacks are nearly identical. SQL injection, whose flows differ more, is the real test, and most models fail it. The random forest falls from 0.95 to 0.43, the MLP from 0.90 to 0.48, logistic regression from 0.81 to 0.57 and the linear SVM from 0.95 to 0.71. Only gradient-boosted trees remain strong (1.00 to 0.90). Supervised detectors largely recognise the attacks they were shown, and the threshold IDS detects nothing in any case."));
c.push(...figure("fig07_unseen_attacks.png", "Fig. 7. Detection rate of each attack type when it was seen in training (hollow, out-of-fold CV) and when it was held out entirely (filled)."));

c.push(h2("5.7 Low-and-slow timing evasion"));
c.push(p("Slowing an attack down exposes the brittleness of the most accurate models (Fig. 8). The random forest’s detection rate collapses from 0.99 to 0.10 at only a 2× slow-down, consistent with its dependence on Fwd IAT Min. Gradient-boosted trees degrade to 0.88 at 5×, 0.28 at 10× and 0.02 at 20×. The MLP holds until 20× (0.89) and then fails. The linear models are the most robust (at least 0.94 at every factor), because a log-scaled linear score shifts only gradually. The conventional threshold IDS behaves in the opposite way: extreme slow-downs push durations outside the benign range, and from 20× onwards it detects 93–94% of attacks. Accuracy and robustness pull in different directions, and the two detector families fail in different regimes."));
c.push(...figure("fig08_timing_evasion.png", "Fig. 8. Detection rate under simulated within-connection slow-down (time features ×k, rates ÷k).", 0.85));

c.push(h2("5.8 False alarms on an unseen day"));
c.push(p("On Monday’s 47,799 benign HTTP flows (Fig. 9, Table 5), the false-positive rate of the linear models more than doubles (to 2.4%, about 1,150 false alerts) and the MLP’s rises to 1.1%. The threshold IDS rises to 1.9%, because its benign profile is also day-specific. The tree ensembles stay remarkably stable: 11 false alerts for the random forest and 61 for gradient-boosted trees over the whole day. Re-calibrating each detector’s alarm threshold on a known-clean day, here a 1% budget on Monday, restores control without losing detection. Thursday recall at the re-calibrated thresholds is 1.000 for both tree ensembles and 0.97–0.99 for the other models."));
c.push(table([2450, 1300, 1350, 1300, 1600, 1350], ["Detector", "FPR Thursday", "FPR Monday", "False alerts (Mon.)", "Recall after re-calibration", "FPR Thu. (re-cal.)"], [
  ["Logistic Regression", "1.08%", "2.44%", "1,168", "0.967", "0.34%"],
  ["Linear SVM", "0.77%", "2.38%", "1,138", "0.984", "0.43%"],
  ["Random Forest", "0.00%", "0.02%", "11", "1.000", "0.54%"],
  ["Gradient-Boosted Trees", "0.03%", "0.13%", "61", "1.000", "0.43%"],
  ["Neural Network (MLP)", "0.43%", "1.08%", "517", "0.988", "0.40%"],
  ["Threshold IDS (conv.)", "0.68%", "1.92%", "920", "0.000", "0.37%"],
], { highlight: (i) => i === 5 }));
c.push(cap("Table 5. False alarms on Thursday’s test split and on an entire unseen benign day (Monday), and Thursday detection after re-calibrating thresholds to a 1% budget on Monday."));
c.push(...figure("fig09_cross_day.png", "Fig. 9. False-positive rate on Thursday’s test split vs. an unseen benign day.", 0.9));

c.push(h2("5.9 Evaluation scope"));
c.push(p("Evaluating on all Thursday traffic (164,179 flows, 1.3% attacks) instead of HTTP flows inflates ranking metrics (Table 6). A rule that knows nothing but “destination port = 80” reaches ROC-AUC 0.947 with 100% recall, and false-positive rates are diluted by tens of thousands of easy non-web flows. The random forest scores about the same in both scopes. Results from the all-traffic setting therefore partly measure web-versus-non-web separation, which is why this project reports the HTTP scope."));
c.push(table([3100, 1050, 1050, 1050, 1150, 1050], ["Setting", "Precision", "Recall", "F1", "ROC-AUC", "FPR"], [
  ["Rule: destination port = 80 (all traffic)", "0.110", "1.000", "0.199", "0.947", "10.67%"],
  ["Random Forest, all traffic, with port", "0.995", "0.979", "0.987", "1.000", "0.01%"],
  ["Random Forest, all traffic, no port", "1.000", "0.970", "0.985", "1.000", "0.00%"],
  ["Random Forest, HTTP flows only (ours)", "1.000", "0.986", "0.993", "1.000", "0.00%"],
], { highlight: (i) => i === 0 }));
c.push(cap("Table 6. Effect of the evaluation scope (held-out test sets)."));

c.push(h2("5.10 AI-assisted detection vs. conventional scanning"));
c.push(p("Table 7 summarises the comparison. On flow data, AI is decisively better than the conventional threshold IDS at detection, at attributing the attack type, and at controlling false alarms on new days. The threshold IDS keeps the conventional advantages of needing no labelled attacks, being fully transparent, and resisting extreme slow-downs. Signature scanners and active scanners such as OWASP ZAP work on payloads and responses rather than flows. They confirm specific vulnerabilities and name them precisely, but they need visibility into the application traffic and a maintained rule base. Flow-based AI and payload-based scanning therefore answer different questions: “is someone attacking us right now?” versus “which flaws does the application have?”"));
c.push(table([2600, 2250, 2250, 2250], ["Dimension", "AI/ML on flows (this work)", "Threshold IDS (conventional)", "Signature / active scanner (ZAP)"], [
  ["Input", "Flow statistics", "Flow statistics", "Payloads, live application"],
  ["Detection F1 (HTTP flows)", "0.998 (S0) / 0.968 (S3)", "0.000", "n/a on flow data"],
  ["Names the attack type", "Partly (macro-F1 0.90 in S3)", "No", "Yes (rule / CWE)"],
  ["Needs labelled attacks", "Yes", "No", "No (needs rules)"],
  ["Unseen attack type (SQLi)", "0.43–0.90 recall", "0.00", "Only if a rule exists"],
  ["Low-and-slow evasion", "Trees fragile; linear robust", "Detects extreme slow-downs", "Unaffected (payload-based)"],
  ["Explainability", "Partial (feature importance)", "Full (metrics that fired)", "Full (rule that fired)"],
  ["Confirms exploitability", "No", "No", "Yes (active scanner)"],
]));
c.push(cap("Table 7. Capabilities of AI-assisted flow detection and conventional approaches (measured values from this study where available)."));
c.push(p("Practical component. The repository includes an AI-assisted assessment prototype and a lab harness for an intentionally vulnerable application (OWASP Juice Shop). In phase A, benign browsing is captured to a PCAP file. In phase B, an OWASP ZAP full active scan is captured while ZAP writes its own report. Both captures are converted to flows with CICFlowMeter and scored by every detector, and a report compares ZAP’s confirmed findings with the AI’s detection rate on ZAP’s attack traffic and its false-positive rate on benign browsing. A column mapper reconciles the CIC-IDS2017 names with both CICFlowMeter implementations, verified against their source code, and converts the Python implementation’s seconds to the microseconds used in training. The harness was validated end to end: test flows exported in both formats reproduced the models’ direct predictions exactly. The live capture must be run in an authorised lab, and results there should be read with E5 in mind, since a new network and flow-meter version shift the statistics."));

c.push(h2("5.11 Computational cost"));
c.push(p("All models are cheap enough for inline use (Table 8). Training takes 0.2–6.4 s on 15,780 flows on two CPU cores, and inference costs 0.004–0.04 ms per flow, so even the random forest scores about 24,000 flows per second."));
c.push(table([3000, 2000, 2200], ["Detector", "Training time (s)", "Inference (ms per flow)"], [
  ["Logistic Regression", "0.29", "0.006"], ["Linear SVM", "0.17", "0.004"], ["Random Forest", "2.97", "0.041"],
  ["Gradient-Boosted Trees", "6.43", "0.010"], ["Neural Network (MLP)", "1.65", "0.006"], ["Threshold IDS (conv.)", "0.03", "0.003"],
]));
c.push(cap("Table 8. Training and inference cost (held-out experiment, 2 CPU cores)."));

// 6 Discussion
c.push(h1("6. Discussion"));
c.push(p([R("Answer to the research question. ", { b: true }), R("AI/ML assists web-attack detection effectively. From flow statistics alone it detects web attacks the conventional threshold approach cannot see at all. In the most conservative setting it still reaches F1 0.97 and attack-type macro-F1 0.90, at negligible computational cost. Its advantage comes from combining many weak signals, since the attacks are ordinary on every metric taken separately. The headline accuracy reported for this dataset, however, overstates real-world performance. About 91% of attack-labelled flows are empty connection fragments, a host fingerprint is learnable, unseen attack types are often missed, timing manipulation defeats the tree ensembles, and false alarms rise on a new day unless thresholds are re-calibrated.")]));
c.push(p([R("Practical implications. ", { b: true }), R("(1) Evaluate on the traffic that matters (here, HTTP flows that carry payloads) and audit for fingerprints before trusting a score. (2) Combine model families: the tree ensembles are accurate and stable across days but fragile to timing manipulation, while linear models and simple statistical thresholds are robust to it. (3) Re-calibrate alarm thresholds per site on known-clean traffic. (4) Use flow-based AI for monitoring and triage, and payload-based tools (WAF rules, active scanners such as ZAP) to confirm and name vulnerabilities.")]));
c.push(p([R("Limitations and threats to validity. ", { b: true }), R("The study uses one dataset and one attack day, and SQLi has only 21 flows (12 with payload), so its per-class numbers are uncertain despite cross-validation. The evasion experiment perturbs features rather than replaying slowed attacks. Delays between separate connections would not change per-flow statistics at all. The live ZAP comparison is provided as a validated harness but must be executed in an authorised lab. Deep sequence models (LSTM, GRU, transformers) were not evaluated because TensorFlow and PyTorch were unavailable in the build environment. The models use a single configuration and seed. CIC-IDS2017 has known construction errors beyond those audited here [2], [3].")]));

// 7 Conclusion
c.push(h1("7. Conclusion and Future Work"));
c.push(p("On the CIC-IDS2017 web-attack day, AI/ML models detected web attacks from flow statistics far better than a conventional threshold IDS, which missed every attack because the attacks are statistically ordinary flows. Gradient-boosted trees and random forests were the strongest and most stable detectors. A validity audit showed that standard evaluations inflate detection and depress attack-type classification, because most attack flows are payload-less artifacts and a TCP fingerprint is learnable. With these controlled, ML still detected real attack requests with F1 0.97 and named the attack type with macro-F1 0.90. Robustness tests showed that unseen attack types, timing evasion and day-to-day drift remain open problems. Future work: re-run the study on a corrected release of the dataset [2], [3]; evaluate sequence models on raw packets; train ensembles that are robust to timing evasion; add per-host aggregation of flows; and complete the live OWASP ZAP comparison with the provided harness."));

// References
c.push(h1("References"));
const refs = [
  "I. Sharafaldin, A. H. Lashkari, and A. A. Ghorbani, “Toward generating a new intrusion detection dataset and intrusion traffic characterization,” in Proc. 4th Int. Conf. on Information Systems Security and Privacy (ICISSP), Portugal, 2018.",
  "G. Engelen, V. Rimmer, and W. Joosen, “Troubleshooting an intrusion detection dataset: the CICIDS2017 case study,” in Proc. IEEE Security and Privacy Workshops (SPW), 2021, doi: 10.1109/SPW53761.2021.00009.",
  "M. Lanvin, P.-F. Gimenez, Y. Han, F. Majorczyk, L. Mé, and É. Totel, “Errors in the CICIDS2017 dataset and the significant differences in detection performances it makes,” in Risks and Security of Internet and Systems (CRiSIS 2022), LNCS vol. 13857, Springer, 2023, doi: 10.1007/978-3-031-31108-6_2.",
  "Z. K. Maseer, R. Yusof, N. Bahaman, S. A. Mostafa, and C. F. M. Foozy, “Benchmarking of machine learning for anomaly based intrusion detection systems in the CICIDS2017 dataset,” IEEE Access, vol. 9, pp. 22351–22370, 2021.",
  "Kurniabudi, D. Stiawan, Darmawijoyo, M. Y. Bin Idris, A. M. Bamhdi, and R. Budiarto, “CICIDS-2017 dataset feature analysis with information gain for anomaly detection,” IEEE Access, vol. 8, pp. 132911–132921, 2020, doi: 10.1109/ACCESS.2020.3009843.",
  "D. Arp, E. Quiring, F. Pendlebury, A. Warnecke, F. Pierazzi, C. Wressnegger, L. Cavallaro, and K. Rieck, “Dos and don’ts of machine learning in computer security,” in Proc. 31st USENIX Security Symposium, 2022.",
  "R. Sommer and V. Paxson, “Outside the closed world: On using machine learning for network intrusion detection,” in Proc. IEEE Symposium on Security and Privacy, 2010.",
  "D. E. Denning, “An intrusion-detection model,” IEEE Transactions on Software Engineering, vol. SE-13, no. 2, 1987.",
  "J. Liang, W. Zhao, and W. Ye, “Anomaly-based web attack detection: A deep learning approach,” in Proc. 6th Int. Conf. on Network, Communication and Computing (ICNCC), ACM, 2017.",
  "F. A. Mereani and J. M. Howe, “Detecting cross-site scripting attacks using machine learning,” Advances in Intelligent Systems and Computing, vol. 723, pp. 200–210, Springer, 2018, doi: 10.1007/978-3-319-74690-6_20.",
];
refs.forEach((r, i) => c.push(new Paragraph({ spacing: { after: 80 }, alignment: AlignmentType.JUSTIFIED,
  children: [new TextRun({ text: `[${i + 1}] `, bold: true, size: 19, font: FONT }), new TextRun({ text: r, size: 19, font: FONT })] })));

// Appendices
c.push(h1("Appendix A. Reproducing the results"));
c.push(p("Code, figures and result tables are in the project repository. With Python 3.10+ and the packages in requirements.txt, run from the src directory: python download_data.py (download, verify and preprocess); python train_eval.py (E1, E2); python unseen_attacks.py (E3); python evasion_timing.py (E4); python cross_day.py (E5); python scope_shortcut.py (E6); python validity_audit.py (E7); python make_figures.py. Alternatively, run_all.py runs the whole pipeline. All randomness is seeded (seed 42), so every number in this paper is reproducible. The lab harness is in src/lab_harness with its own README."));
c.push(h1("Appendix B. The 66 flow features"));
c.push(table([2300, 7060], ["Group", "Features"], [
  ["Volume", "Flow Duration; Total Fwd/Backward Packets; Total Length of Fwd/Bwd Packets; Subflow Fwd/Bwd Packets and Bytes; act_data_pkt_fwd"],
  ["Packet lengths", "Fwd/Bwd Packet Length Max, Min, Mean, Std; Min/Max Packet Length; Packet Length Mean, Std, Variance; Average Packet Size; Avg Fwd/Bwd Segment Size"],
  ["Timing", "Flow IAT Mean, Std, Max, Min; Fwd IAT Total, Mean, Std, Max, Min; Bwd IAT Total, Mean, Std, Max, Min; Active Mean, Std, Max, Min; Idle Mean, Std, Max, Min"],
  ["Rates", "Flow Bytes/s; Flow Packets/s; Fwd Packets/s; Bwd Packets/s; Down/Up Ratio"],
  ["TCP flags", "Fwd PSH Flags; FIN, SYN, RST, PSH, ACK, URG, ECE Flag Count"],
  ["Headers & TCP stack", "Fwd/Bwd Header Length; Init_Win_bytes_forward; Init_Win_bytes_backward; min_seg_size_forward"],
]));
c.push(cap("Table B1. Model input features by group (CIC-IDS2017 names)."));

// ---------------------------------------------------------------- document
const doc = new Document({
  creator: "CSCI561 project", title: "AI-Assisted Web Application Penetration Testing (CIC-IDS2017)",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 27, bold: true, color: HEAD }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 23, bold: true, color: "2E5A80" }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: "bullet", text: "•",
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 18, font: FONT })] })] }) },
    children: c,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Wrote", OUT, buf.length, "bytes"); });
