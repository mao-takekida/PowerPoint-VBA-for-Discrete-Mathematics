グラフを作成する時に使うマクロ.

- `make_edge` 
    - `MakeEdgeUndirected1ToN` : 一つ目に選択した頂点から二つ目以降に選択した頂点全てに無向辺を張る.
    - `MakeEdgeDirected1ToN` : 一つ目に選択した頂点から二つ目以降に選択した頂点全てに有向辺を張る.
- `select_only`
    - `SelectOnlyEdges` : 現在選択中のオブジェクトのうち, 辺だけを選択した状態にする.
    - `SelectOnlyOvals` : 現在選択中のオブジェクトのうち, 楕円(円)だけを選択した状態にする.
    - `SelectOnlyDodecagons` : 現在選択中のオブジェクトのうち, 十角形だけを選択した状態にする.
    - `SelectOnlyTextBoxes` : 現在選択中のオブジェクトのうち, テキストボックスだけを選択した状態にする.
- `assign_number`
    - `AssignNumbers` : テキストフレームに選択順で番号を振る.
- `assign_weight`
    - `AssignWeightsRandom` : 選択中の辺にランダムな重みを振る.
