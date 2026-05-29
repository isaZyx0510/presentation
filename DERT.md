```mermaid
flowchart TD
    CFG["cfg config<br/>N_rx, N_tx, L_eff, N_sc<br/>N_obs = len observedSymbols<br/>d_model = transformerDModel<br/>q = max L_eff and detrNumQueries<br/>mu_mode = detrMuMode, default scaled<br/>gain_mode = gainMode, default free-running"]:::cfg

    X["Input x<br/>B x 2N_rx x N_sc x N_obs"]:::io
    CHECK{"Shape check<br/>expect B x 2N_rx x N_sc x N_obs"}:::condition
    ERR["raise ValueError<br/>wrong input shape"]:::error

    STEM["CNN stem<br/>Conv2d 2N_rx to d_model<br/>GELU<br/>Conv2d d_model to d_model<br/>GELU"]:::module
    FMAP["fmap<br/>B x d_model x N_sc x N_obs"]:::process

    PERMUTE["permute<br/>B x d_model x N_sc x N_obs<br/>to B x N_sc x N_obs x d_model"]:::process
    RESHAPE_TOKENS["reshape to tokens<br/>T = N_sc x N_obs<br/>fseq: B x T x d_model"]:::process

    POS_K["pos_k learnable parameter<br/>N_sc x d_model"]:::process
    POS_N["pos_n learnable parameter<br/>N_obs x d_model"]:::process
    POS_ADD["broadcast add<br/>pos_k plus pos_n<br/>N_sc x N_obs x d_model"]:::process
    POS_RESHAPE["reshape position<br/>pos: T x d_model"]:::process
    POS_UNSQUEEZE["unsqueeze batch dim<br/>pos: 1 x T x d_model"]:::process
    ADD_POS["add position<br/>encoder_input: B x T x d_model"]:::process

    ENC["Transformer Encoder<br/>self attention over T tokens"]:::module
    MEMORY["memory<br/>B x T x d_model"]:::process

    QUERY_PARAM["learnable query parameters<br/>q x d_model"]:::process
    QUERY_EXPAND["unsqueeze and expand<br/>queries: B x q x d_model"]:::process

    DEC["Transformer Decoder<br/>query self attention<br/>cross attention from queries to memory"]:::module
    DECODED["decoded<br/>B x q x d_model"]:::process

    ALPHA_HEAD["alpha_head<br/>Linear d_model to 1<br/>squeeze last dim<br/>softplus"]:::module
    ALPHA_ALL["alpha_all<br/>B x q"]:::process

    OBJ_HEAD["obj_head<br/>Linear d_model to 1<br/>squeeze last dim<br/>sigmoid"]:::module
    PATH_PROB_ALL["path_prob_all<br/>B x q"]:::process

    MU_MODE{"mu_mode"}:::condition
    MU_FREE["free branch<br/>mu_head: B x q x 1<br/>squeeze to B x q<br/>sigmoid and scale to range"]:::module
    MU_RAW_FREE["mu_raw_all<br/>B x q"]:::process

    MU_MEAN["scaled or grid branch<br/>mean over query dim<br/>cross_path_stats: B x d_model"]:::process
    S_HEAD["s_head MLP<br/>B x d_model to B x 1<br/>sigmoid and scale"]:::module
    EXTENT["extent<br/>B x 1"]:::process

    MU_SCALED["scaled branch<br/>mu_head: B x q x 1<br/>squeeze to B x q<br/>sigmoid"]:::module
    MU_GRID["grid branch<br/>mu_head: B x q x num_mu_bins<br/>softmax over bins<br/>weighted sum with mu_dict"]:::module
    MU_FRAC["frac<br/>B x q"]:::process
    MU_COMBINE["combine frac and extent<br/>mu_raw_all = mu_min plus frac times extent range<br/>B x q"]:::process
    MU_RAW_ALL["mu_raw_all<br/>B x q"]:::process

    RHO_HEAD["rho_head<br/>Linear d_model to 1<br/>squeeze last dim<br/>sigmoid and scale to rho range"]:::module
    RHO_RAW_ALL["rho_raw_all<br/>B x q"]:::process

    GAIN_MODE{"gain_mode"}:::condition
    G_FREE_HEAD["free-running branch<br/>g_head: B x q x 2N_rxN_tx"]:::module
    G_FREE_VIEW["reshape and view_as_complex<br/>B x q x N_rx x N_tx x 2<br/>to complex B x q x N_rx x N_tx"]:::process

    G_A_HEAD["rank1 branch a_head<br/>B x q x 2N_rx<br/>reshape to B x q x N_rx x 2<br/>view_as_complex"]:::module
    G_B_HEAD["rank1 branch b_head<br/>B x q x 2N_tx<br/>reshape to B x q x N_tx x 2<br/>view_as_complex"]:::module
    G_OUTER["complex outer product<br/>a times conjugate b<br/>g_raw: B x q x N_rx x N_tx"]:::process

    G_RAW["g_raw complex<br/>B x q x N_rx x N_tx"]:::process
    G_FRO["Frobenius norm<br/>fro: B x q<br/>clamp min 1e-8"]:::process
    G_NORM["normalize and scale<br/>gains_all = g_raw divided by fro times alpha_all<br/>B x q x N_rx x N_tx"]:::process

    SCORE["score = alpha_all times path_prob_all<br/>B x q"]:::process
    TOPK["topk over query dim<br/>top_idx: B x L_eff"]:::process

    GATHER_MU["gather mu_raw_all using top_idx<br/>mu_raw: B x L_eff"]:::process
    GATHER_RHO["gather rho_raw_all using top_idx<br/>rho_raw: B x L_eff"]:::process
    GATHER_ALPHA["gather alpha_all using top_idx<br/>alpha: B x L_eff"]:::process
    GATHER_PROB["gather path_prob_all using top_idx<br/>path_prob: B x L_eff"]:::process
    EXPAND_TOPK["expand top_idx<br/>B x L_eff x N_rx x N_tx"]:::process
    GATHER_G["gather gains_all using expanded top_idx<br/>G_hat: B x L_eff x N_rx x N_tx complex"]:::process

    GAUGE_MU["gauge_fix_mu<br/>mu0 = max mu_raw<br/>mu_rel = mu_raw - mu0"]:::gauge
    MU0["mu0<br/>B x 1"]:::process
    MU_REL["mu_rel<br/>B x L_eff"]:::process

    GAUGE_RHO["gauge_fix_rho<br/>rho0 = mean rho_raw<br/>rho_rel = rho_raw - rho0"]:::gauge
    RHO0["rho0<br/>B x 1"]:::process
    RHO_REL["rho_rel<br/>B x L_eff"]:::process

    DELTA_POOL["AdaptiveAvgPool2d<br/>B x d_model x 1 x 1"]:::process
    DELTA_FLAT["flatten<br/>B x d_model"]:::process
    DELTA_MLP["delta_mlp<br/>MLP B x d_model to B x N_rx<br/>sigmoid and scale to delta range"]:::module
    DELTA_RAW["delta_raw<br/>B x N_rx"]:::process
    GAUGE_DELTA["gauge_fix_delta<br/>add or set reference antenna delta0 = 0<br/>return rx relative offsets"]:::gauge
    DELTA_RX["delta_rx<br/>B x N_rx"]:::process

    MIX_PARAM["mix_lambda parameter<br/>1"]:::process
    MIX_SIGMOID["sigmoid<br/>mix_lambda: 1"]:::process

    OUT_THETA["Output theta dict<br/>mu0: B x 1<br/>rho0: B x 1<br/>delta_rx: B x N_rx<br/>mu_rel: B x L_eff<br/>rho_rel: B x L_eff"]:::io
    OUT_G["Output G_hat<br/>complex<br/>B x L_eff x N_rx x N_tx"]:::io
    OUT_ALPHA["Output alpha<br/>B x L_eff"]:::io
    OUT_PROB["Output path_prob<br/>B x L_eff"]:::io
    OUT_MIX["Output mix_lambda<br/>1"]:::io

    NOTE_UNUSED["Note<br/>_sincos_pos is defined but not used in this forward path<br/>pack_slope_theta is not used<br/>model returns dict instead of packed_theta"]:::note
    NOTE_NO_CONCAT["Note<br/>No concat operation in this forward path<br/>Main reshape operations are token reshape and real-imag to complex"]:::note

    CFG -.-> CHECK
    X --> CHECK
    CHECK -- invalid --> ERR
    CHECK -- valid --> STEM

    STEM --> FMAP
    FMAP --> PERMUTE
    PERMUTE --> RESHAPE_TOKENS

    POS_K --> POS_ADD
    POS_N --> POS_ADD
    POS_ADD --> POS_RESHAPE
    POS_RESHAPE --> POS_UNSQUEEZE
    RESHAPE_TOKENS --> ADD_POS
    POS_UNSQUEEZE --> ADD_POS

    ADD_POS --> ENC
    ENC --> MEMORY

    QUERY_PARAM --> QUERY_EXPAND

    MEMORY --> DEC
    QUERY_EXPAND --> DEC
    DEC --> DECODED

    DECODED --> ALPHA_HEAD
    ALPHA_HEAD --> ALPHA_ALL

    DECODED --> OBJ_HEAD
    OBJ_HEAD --> PATH_PROB_ALL

    CFG -.-> MU_MODE
    DECODED --> MU_MODE
    MU_MODE -- free --> MU_FREE
    MU_FREE --> MU_RAW_FREE
    MU_RAW_FREE --> MU_RAW_ALL

    MU_MODE -- scaled --> MU_MEAN
    MU_MODE -- grid --> MU_MEAN
    MU_MEAN --> S_HEAD
    S_HEAD --> EXTENT

    MU_MODE -- scaled --> MU_SCALED
    MU_MODE -- grid --> MU_GRID
    MU_SCALED --> MU_FRAC
    MU_GRID --> MU_FRAC
    MU_FRAC --> MU_COMBINE
    EXTENT --> MU_COMBINE
    MU_COMBINE --> MU_RAW_ALL

    DECODED --> RHO_HEAD
    RHO_HEAD --> RHO_RAW_ALL

    CFG -.-> GAIN_MODE
    DECODED --> GAIN_MODE
    GAIN_MODE -- free-running --> G_FREE_HEAD
    G_FREE_HEAD --> G_FREE_VIEW
    G_FREE_VIEW --> G_RAW

    GAIN_MODE -- rank1 --> G_A_HEAD
    GAIN_MODE -- rank1 --> G_B_HEAD
    G_A_HEAD --> G_OUTER
    G_B_HEAD --> G_OUTER
    G_OUTER --> G_RAW

    G_RAW --> G_FRO
    G_RAW --> G_NORM
    G_FRO --> G_NORM
    ALPHA_ALL --> G_NORM

    ALPHA_ALL --> SCORE
    PATH_PROB_ALL --> SCORE
    SCORE --> TOPK

    MU_RAW_ALL --> GATHER_MU
    TOPK --> GATHER_MU

    RHO_RAW_ALL --> GATHER_RHO
    TOPK --> GATHER_RHO

    ALPHA_ALL --> GATHER_ALPHA
    TOPK --> GATHER_ALPHA

    PATH_PROB_ALL --> GATHER_PROB
    TOPK --> GATHER_PROB

    TOPK --> EXPAND_TOPK
    G_NORM --> GATHER_G
    EXPAND_TOPK --> GATHER_G

    GATHER_MU --> GAUGE_MU
    GAUGE_MU --> MU0
    GAUGE_MU --> MU_REL

    GATHER_RHO --> GAUGE_RHO
    GAUGE_RHO --> RHO0
    GAUGE_RHO --> RHO_REL

    FMAP --> DELTA_POOL
    DELTA_POOL --> DELTA_FLAT
    DELTA_FLAT --> DELTA_MLP
    DELTA_MLP --> DELTA_RAW
    DELTA_RAW --> GAUGE_DELTA
    GAUGE_DELTA --> DELTA_RX

    MIX_PARAM --> MIX_SIGMOID

    MU0 --> OUT_THETA
    RHO0 --> OUT_THETA
    DELTA_RX --> OUT_THETA
    MU_REL --> OUT_THETA
    RHO_REL --> OUT_THETA

    GATHER_G --> OUT_G
    GATHER_ALPHA --> OUT_ALPHA
    GATHER_PROB --> OUT_PROB
    MIX_SIGMOID --> OUT_MIX

    classDef io fill:#B7E4C7,stroke:#1B4332,stroke-width:2px,color:#111827;
    classDef process fill:#A9D6E5,stroke:#014F86,stroke-width:2px,color:#111827;
    classDef module fill:#FFE08A,stroke:#B7791F,stroke-width:2px,stroke-dasharray:5 5,color:#111827;
    classDef condition fill:#F8C4C4,stroke:#9B2226,stroke-width:2px,color:#111827;
    classDef error fill:#F4A6A6,stroke:#7F1D1D,stroke-width:2px,color:#111827;
    classDef gauge fill:#F6B26B,stroke:#B45309,stroke-width:2px,color:#111827;
    classDef cfg fill:#CDB4DB,stroke:#6D28D9,stroke-width:2px,color:#111827;
    classDef note fill:#D8B4E2,stroke:#7E22CE,stroke-width:2px,color:#111827;