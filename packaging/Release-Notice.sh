#!/usr/bin/env bash
# Tells the open issues a release covers that the changes are out: one comment per issue and release, in the issue's language.
# Usage: Release-Notice.sh <version> [dry-run]   (needs gh, jq, GH_TOKEN and GITHUB_REPOSITORY; dry-run only prints)
# Run by .github/workflows/release-notice.yml after "Publish release"; see packaging/README.md.
#
# Which issues: the "(#N)" / "(#N, #M)" at the end of entries in the release notes (the CHANGELOG section). Only issues that are
# still open (closing one first opts it out) and not pull requests; issues the owner opened (transcribed from forums) get one
# too, as a record of the release. The issue is left open: whether it is done is decided by hand.
# Each comment ends with a hidden marker naming the release, so running again for the same release skips issues already told.
set -uo pipefail

version=${1:?usage: Release-Notice.sh <version> [dry-run]}
dryRun=false
[[ "${2:-}" == dry-run ]] && dryRun=true
[[ "$version" =~ ^[0-9]{4}\.[0-9]{2}\.[0-9]{2}(\.[0-9]+)?$ ]] || { echo "Not a version: $version"; exit 1; }
repo=${GITHUB_REPOSITORY:?}
bot="github-actions[bot]"
summary=${GITHUB_STEP_SUMMARY:-/dev/null}

# test builds have no GitHub release: nothing to tell
if ! release=$(gh release view "$version" --repo "$repo" --json tagName,isPrerelease,isDraft,url,body 2>/dev/null); then
    echo "No release $version: nothing to do"
    exit 0
fi
if [[ $(jq -r '.isPrerelease or .isDraft' <<<"$release") == true ]]; then
    echo "Release $version is a draft or a prerelease: nothing to do"
    exit 0
fi
url=$(jq -r .url <<<"$release")
marker="<!-- altrun-release-notice:stable $version -->"

numbers=$(jq -r .body <<<"$release" | grep -oE '\(#[0-9]+(, #[0-9]+)*\)' | grep -oE '[0-9]+' | sort -un || true)
echo "Release $version ($url), issues in the notes: $(tr "\n" " " <<<"${numbers:-none}")"
echo "### Release notice for $version$([[ $dryRun == true ]] && echo ' (dry run)')" >>"$summary"
[[ -z "$numbers" ]] && { echo "No issues referenced" >>"$summary"; exit 0; }

failed=0
for n in $numbers; do
    if ! issue=$(gh api "repos/$repo/issues/$n" 2>&1); then
        echo "#$n: could not read the issue: $issue"; echo "- #$n: **failed** to read" >>"$summary"; failed=$((failed + 1)); continue
    fi
    skip=$(jq -r '
        if has("pull_request") then "a pull request"
        elif .state != "open" then "closed"
        else "" end' <<<"$issue")
    if [[ -n "$skip" ]]; then
        echo "#$n: skipped ($skip)"; echo "- #$n: skipped ($skip)" >>"$summary"; continue
    fi
    if ! told=$(gh api --paginate "repos/$repo/issues/$n/comments" \
            --jq ".[] | select(.user.login == \"$bot\") | .body | contains(\"$marker\")" 2>&1); then
        echo "#$n: could not read the comments: $told"; echo "- #$n: **failed** to read comments" >>"$summary"; failed=$((failed + 1)); continue
    fi
    if grep -qx true <<<"$told"; then
        echo "#$n: already told"; echo "- #$n: already told" >>"$summary"; continue
    fi

    # Chinese when the title or the text has Chinese characters; the user's own text is only read here, never put in a command
    if [[ $(jq -r '((.title // "") + " " + (.body // "")) | test("\\p{Han}")' <<<"$issue") == true ]]; then
        text="相关改动已包含在正式版 [ALTRun $version]($url) 中. 可以在托盘菜单 → 检查更新 里升级, 也可以从发布页下载. 如果还有问题, 请在这里回复."
        lang=zh
    else
        text="The changes for this issue are in [ALTRun $version]($url). Update from the tray menu → Check for Updates, or download it from the release page. If something is still not right, please reply here."
        lang=en
    fi
    body="$text"$'\n\n'"$marker"
    if [[ $dryRun == true ]]; then
        echo "#$n: would comment ($lang): $text"; echo "- #$n: would comment ($lang)" >>"$summary"; continue
    fi
    if out=$(gh issue comment "$n" --repo "$repo" --body "$body" 2>&1); then
        echo "#$n: commented ($lang) $out"; echo "- #$n: commented ($lang)" >>"$summary"
    else
        echo "#$n: comment failed: $out"; echo "- #$n: **comment failed**" >>"$summary"; failed=$((failed + 1))
    fi
done

if ((failed > 0)); then
    echo "$failed issue(s) failed; run Release notice again (Actions → Release notice → Run workflow) to retry them"
    exit 1
fi
