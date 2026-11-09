#!/usr/bin/env bash
# Branch-only measurement of a full CPU build from upstream source.
set -euo pipefail

readonly ort_commit=6af1a932e2286d5c58166085d24c68783033edee
readonly python=/opt/python/cp312-cp312/bin/python3.12
readonly wheel=/work/out/onnxruntime-1.31.1-cp312-cp312-manylinux_2_28_x86_64.whl
readonly out=/work/out
test "$(uname -m)" = x86_64

record_resources() {
  date -u --iso-8601=seconds
  printf 'elapsed_seconds=%s\n' "$SECONDS"
  printf 'cpus=%s\n' "$(nproc)"
  awk '/MemTotal:/ {print}' /proc/meminfo
  if [[ -r /sys/fs/cgroup/memory.peak ]]; then
    printf 'cgroup_peak_memory_bytes='
    cat /sys/fs/cgroup/memory.peak
  fi
  df -k /work
  du -sk /work/ort-source /work/ort-build /work/build-venv 2>/dev/null || true
}

case "$1" in
  build)
    trap 'record_resources > "$out/resources-end.txt"' EXIT
    record_resources > "$out/resources-start.txt"
    git init -q ort-source
    git -C ort-source remote add origin https://github.com/microsoft/onnxruntime.git
    git -C ort-source fetch --no-tags --depth 1 origin "$ort_commit"
    git -C ort-source checkout --detach FETCH_HEAD
    test "$(git -C ort-source rev-parse HEAD)" = "$ort_commit"
    git -C ort-source submodule update --init --recursive --depth 1
    git -C ort-source rev-parse HEAD > "$out/ort.commit"
    git -C ort-source submodule status --recursive > "$out/ort.submodules"
    if grep -Eq '^[+-U]' "$out/ort.submodules"; then
      echo "An ORT submodule does not match its recorded commit" >&2
      exit 1
    fi
    cp ort-source/cmake/deps.txt "$out/ort-cmake-deps.txt"

    "$python" -m venv /work/build-venv
    export PATH="/work/build-venv/bin:$PATH"
    # Hashes select these exact wheel inputs for Linux x86-64 and CPython 3.12.
    cat > "$out/build-requirements.txt" <<'LOCK'
cmake==3.31.10 --hash=sha256:2f766bb46367e5e0559fa33184653754bce044583a06014dcaebf8e6dff8a1f1
ninja==1.13.0 --hash=sha256:fb46acf6b93b8dd0322adc3a4945452a4e774b75b91293bafcc7b7f8e6517dfa
setuptools==84.1.0 --hash=sha256:5231689c6ac8156cd155cfb98c3fd6fb252c3e7e9bd3d78b64d3c55f132624c9
wheel==0.46.3 --hash=sha256:4b399d56c9d9338230118d705d9737a2a468ccca63d5e813e2a4fc7815d8bc4d
packaging==26.3 --hash=sha256:d7193f7c8e4e93f444fde0262bf90af30e16fa0ad0ad44cb553c87339b23cd1c
numpy==2.4.4 --hash=sha256:81f4a14bee47aec54f883e0cad2d73986640c1590eb9bfaaba7ad17394481e6e
flatbuffers==25.12.19 --hash=sha256:7634f50c427838bb021c2d66a3d1168e9d199b0607e6329399f04846d42e20b4
protobuf==7.34.1 --hash=sha256:8ff40ce8cd688f7265326b38d5a1bed9bfdf5e6723d49961432f83e21d5713e4
LOCK
    python -m pip --isolated install --index-url https://pypi.org/simple \
      --only-binary :all: --no-deps --no-cache-dir --require-hashes -r "$out/build-requirements.txt"
    {
      python --version
      python -m pip --version
      cmake --version
      ninja --version
      gcc --version
      auditwheel --version
    } > "$out/toolchain.txt"
    export AUDITWHEEL_PLAT=manylinux_2_28_x86_64
    export ORT_DISABLE_TELEMETRY=1
    args=(
      --build_dir /work/ort-build --config Release --update --build
      --build_wheel --build_shared_lib --skip_tests --skip_submodule_sync
      --allow_running_as_root --cmake_generator Ninja --parallel 3 --no_telemetry
      --cmake_extra_defines onnxruntime_DISABLE_DEVICE_DISCOVERY=ON onnxruntime_BUILD_UNIT_TESTS=OFF
    )
    printf '%s\n' "${args[@]}" > "$out/ort-build-arguments.txt"
    (cd ort-source && python tools/ci_build/build.py "${args[@]}")
    cache=/work/ort-build/Release/CMakeCache.txt
    grep -E '^onnxruntime_(DISABLE_DEVICE_DISCOVERY|USE_TELEMETRY|BUILD_UNIT_TESTS):' "$cache" \
      > "$out/ort-cmake-options.txt"
    grep -Eq '^onnxruntime_DISABLE_DEVICE_DISCOVERY:[^=]+=ON$' "$out/ort-cmake-options.txt"
    grep -Eq '^onnxruntime_USE_TELEMETRY:[^=]+=OFF$' "$out/ort-cmake-options.txt"
    grep -Eq '^onnxruntime_BUILD_UNIT_TESTS:[^=]+=OFF$' "$out/ort-cmake-options.txt"
    cp "/work/ort-build/Release/dist/$(basename "$wheel")" "$wheel"
    auditwheel show "$wheel" | tee "$out/auditwheel.txt"
    (cd "$out" && sha256sum "$(basename "$wheel")") > "$out/SHA256SUMS.txt"
    python -m pip --isolated install --no-index --no-deps "$wheel"
    python -m pip check
    ;;
  probe)
    export PATH="/work/build-venv/bin:$PATH"
    export ORT_DISABLE_TELEMETRY=1
    (cd "$out" && sha256sum --check SHA256SUMS.txt)
    ln -s /docwen/models /work/probe/models
    sha256sum /docwen/models/rapidocr/ch_PP-OCRv4_det_infer.onnx > "$out/ocr-model.sha256"
    python /work/probe/tests/support/native/verify_ort_optout.py | tee "$out/native-probe.json"
    ;;
  *)
    echo 'usage: probe_ort_linux_wheel.sh build|probe' >&2
    exit 2
    ;;
esac
