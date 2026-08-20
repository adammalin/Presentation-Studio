#!/bin/zsh
set -euo pipefail

readonly PRESENTATION_STUDIO_REPOSITORY="adammalin/Presentation-Studio"
readonly PRESENTATION_STUDIO_BRANCH="codex/web-slide-design-engine"
readonly PRESENTATION_STUDIO_NODE_VERSION="v22.13.0"

INSTALL_ROOT="${PRESENTATION_STUDIO_INSTALL_DIR:-${HOME}/Applications/Presentation Studio}"
APP_DIR="${INSTALL_ROOT}/app"
RUNTIME_ROOT="${INSTALL_ROOT}/runtime"
RUNTIME_NODE_DIR="${RUNTIME_ROOT}/node"
STAGING_DIR="${INSTALL_ROOT}/.app-staging-$$"
PREVIOUS_DIR="${INSTALL_ROOT}/app.previous"
TEMP_DIR="$(mktemp -d "${TMPDIR:-/tmp}/presentation-studio-install.XXXXXX")"

cleanup() {
  if [[ -d "${STAGING_DIR}" ]]; then
    rm -rf -- "${STAGING_DIR}"
  fi
  if [[ -d "${TEMP_DIR}" ]]; then
    rm -rf -- "${TEMP_DIR}"
  fi
}
trap cleanup EXIT

fail() {
  print -u2 "Presentation Studio installation failed: $1"
  exit 1
}

stop_managed_presentation_studio() {
  local runtime_descriptor managed_pid command_line attempt
  runtime_descriptor="${HOME}/Library/Application Support/Presentation Studio/mcp-runtime.json"
  [[ -f "${runtime_descriptor}" ]] || return 0
  managed_pid="$(node -e 'const fs=require("node:fs"); try { const value=JSON.parse(fs.readFileSync(process.argv[1], "utf8")); if (Number.isSafeInteger(value.pid) && value.pid > 1) process.stdout.write(String(value.pid)); } catch {}' "${runtime_descriptor}")"
  [[ -n "${managed_pid}" ]] || return 0
  kill -0 "${managed_pid}" >/dev/null 2>&1 || return 0
  command_line="$(ps -p "${managed_pid}" -o command= 2>/dev/null || true)"
  if [[ "${command_line}" != *"${INSTALL_ROOT}/"*"/node_modules/electron/"* ]]; then
    print "Leaving unrelated running process ${managed_pid} untouched; its executable is outside the managed Presentation Studio install."
    return 0
  fi
  print "Closing the running managed Presentation Studio before activation..."
  kill -TERM "${managed_pid}" >/dev/null 2>&1 || true
  for attempt in {1..80}; do
    kill -0 "${managed_pid}" >/dev/null 2>&1 || return 0
    sleep 0.1
  done
  print "The managed app did not close within 8 seconds; stopping that exact verified process so the update cannot mix runtime versions."
  kill -KILL "${managed_pid}" >/dev/null 2>&1 || true
}

if [[ -z "${INSTALL_ROOT}" || "${INSTALL_ROOT}" == "/" || "${INSTALL_ROOT}" == "${HOME}" ]]; then
  fail "the install location is not safe. Set PRESENTATION_STUDIO_INSTALL_DIR to a dedicated folder."
fi

for command_name in curl shasum tar; do
  command -v "${command_name}" >/dev/null 2>&1 || fail "${command_name} is required by macOS but was not found."
done

if [[ -d "${APP_DIR}" && ! -f "${APP_DIR}/.presentation-studio-managed-install" ]]; then
  fail "${APP_DIR} already exists and is not managed by this installer. Move it or choose another PRESENTATION_STUDIO_INSTALL_DIR."
fi

mkdir -p "${INSTALL_ROOT}" "${RUNTIME_ROOT}"

node_is_compatible() {
  command -v node >/dev/null 2>&1 &&
    command -v npm >/dev/null 2>&1 &&
    node -e 'const [major, minor] = process.versions.node.split(".").map(Number); process.exit(major > 22 || (major === 22 && minor >= 13) ? 0 : 1)' >/dev/null 2>&1
}

install_portable_node() {
  local machine_architecture node_architecture archive_name archive_url checksums_url expected_checksum actual_checksum extracted_dir
  machine_architecture="$(uname -m)"
  case "${machine_architecture}" in
    arm64) node_architecture="arm64" ;;
    x86_64) node_architecture="x64" ;;
    *) fail "unsupported macOS architecture ${machine_architecture}." ;;
  esac

  archive_name="node-${PRESENTATION_STUDIO_NODE_VERSION}-darwin-${node_architecture}.tar.gz"
  archive_url="https://nodejs.org/dist/${PRESENTATION_STUDIO_NODE_VERSION}/${archive_name}"
  checksums_url="https://nodejs.org/dist/${PRESENTATION_STUDIO_NODE_VERSION}/SHASUMS256.txt"

  print "Downloading the portable Node.js prerequisite..."
  curl --fail --location --silent --show-error "${archive_url}" --output "${TEMP_DIR}/${archive_name}"
  curl --fail --location --silent --show-error "${checksums_url}" --output "${TEMP_DIR}/SHASUMS256.txt"
  expected_checksum="$(awk -v archive="${archive_name}" '$2 == archive { print $1 }' "${TEMP_DIR}/SHASUMS256.txt")"
  [[ -n "${expected_checksum}" ]] || fail "Node.js did not publish a checksum for ${archive_name}."
  actual_checksum="$(shasum -a 256 "${TEMP_DIR}/${archive_name}" | awk '{ print $1 }')"
  [[ "${actual_checksum}" == "${expected_checksum}" ]] || fail "the Node.js archive checksum did not match the official manifest."

  mkdir -p "${TEMP_DIR}/node"
  tar -xzf "${TEMP_DIR}/${archive_name}" -C "${TEMP_DIR}/node"
  extracted_dir="${TEMP_DIR}/node/node-${PRESENTATION_STUDIO_NODE_VERSION}-darwin-${node_architecture}"
  [[ -x "${extracted_dir}/bin/node" ]] || fail "the downloaded Node.js runtime is incomplete."

  if [[ -d "${RUNTIME_NODE_DIR}" ]]; then
    rm -rf -- "${RUNTIME_NODE_DIR}"
  fi
  mv "${extracted_dir}" "${RUNTIME_NODE_DIR}"
  export PATH="${RUNTIME_NODE_DIR}/bin:${PATH}"
}

if node_is_compatible; then
  print "Using Node.js $(node --version) from the current system."
else
  install_portable_node
fi

node_is_compatible || fail "Node.js ${PRESENTATION_STUDIO_NODE_VERSION} or newer could not be prepared."

SOURCE_ARCHIVE_URL="${PRESENTATION_STUDIO_SOURCE_ARCHIVE_URL:-https://github.com/${PRESENTATION_STUDIO_REPOSITORY}/archive/refs/heads/${PRESENTATION_STUDIO_BRANCH}.zip}"
print "Downloading the latest Presentation Studio 0.3.1 source..."
curl --fail --location --silent --show-error "${SOURCE_ARCHIVE_URL}" --output "${TEMP_DIR}/presentation-studio.zip"
mkdir -p "${TEMP_DIR}/source"
ditto -x -k "${TEMP_DIR}/presentation-studio.zip" "${TEMP_DIR}/source"

SOURCE_DIR="$(find "${TEMP_DIR}/source" -mindepth 1 -maxdepth 1 -type d -print -quit)"
[[ -n "${SOURCE_DIR}" && -f "${SOURCE_DIR}/package-lock.json" ]] || fail "the downloaded source archive is not a Presentation Studio release."
mv "${SOURCE_DIR}" "${STAGING_DIR}"

print "Installing locked dependencies and verifying the staged application..."
cd "${STAGING_DIR}"
npm ci
npm test
npm run check:data-safety
npm run build
print "managed by scripts/install-macos.sh" > .presentation-studio-managed-install

stop_managed_presentation_studio

if [[ -d "${PREVIOUS_DIR}" ]]; then
  [[ -f "${PREVIOUS_DIR}/.presentation-studio-managed-install" ]] || fail "${PREVIOUS_DIR} is not a managed backup and will not be replaced."
  rm -rf -- "${PREVIOUS_DIR}"
fi
if [[ -d "${APP_DIR}" ]]; then
  mv "${APP_DIR}" "${PREVIOUS_DIR}"
fi
mv "${STAGING_DIR}" "${APP_DIR}"

LAUNCHER="${INSTALL_ROOT}/Launch Presentation Studio.command"
cat > "${LAUNCHER}" <<'LAUNCH_SCRIPT'
#!/bin/zsh
set -euo pipefail
INSTALL_ROOT=${0:A:h}
APP_DIR="${INSTALL_ROOT}/app"
LOCAL_NODE="${INSTALL_ROOT}/runtime/node/bin"
if [[ -x "${LOCAL_NODE}/node" ]]; then
  export PATH="${LOCAL_NODE}:${PATH}"
fi
cd "${APP_DIR}"
npm start
LAUNCH_SCRIPT
chmod 755 "${LAUNCHER}"

print ""
print "Presentation Studio 0.3.1 installed successfully."
print "Install location: ${APP_DIR}"
print "Launcher: ${LAUNCHER}"
print "MCP configuration command: node \"${APP_DIR}/scripts/configure-mcp.mjs\""

if [[ -d "${HOME}/.codex" ]]; then
  print "Configuring the installed Presentation Studio MCP server in Codex..."
  node "${APP_DIR}/scripts/configure-mcp.mjs" --codex "${HOME}/.codex/config.toml"
  print "Restart Codex after installation so it reloads the MCP server list."
fi

if [[ "${PRESENTATION_STUDIO_NO_LAUNCH:-0}" != "1" ]]; then
  print "Starting Presentation Studio..."
  nohup "${LAUNCHER}" > "${INSTALL_ROOT}/presentation-studio.log" 2>&1 &
fi
