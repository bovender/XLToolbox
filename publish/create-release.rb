#! /usr/bin/env ruby
# Creates an XL Toolbox NG release on GitHub
# (c) 2016 Daniel Kraus (bovender)

require 'octokit'

http://stackoverflow.com/a/11868440/270712
branch = `git rev-parse --abbrev-ref HEAD`
if branch != 'master'
  puts "Not on master. Please check out master branch first."
  exit 1
end

client = Octokit::Client.new(netrc: true)
repo = 'bovender/xltoolbox'
tag = `git describe`
tag = 'v7.0.0-beta.4'
msg = `git tag -n9 -l #{tag} | sed -r '1,2d; s/^\\s{4}//'` +
  "\nDownload count for this release: " +
  "![Downloads of #{tag}](https://img.shields.io/github/downloads/bovender/xltoolbox/#{tag}/total.svg?maxAge=2592000)"

version = tag.sub(/^v/, '')
prerelease = !!(tag =~ /alpha|beta/) # http://stackoverflow.com/a/7365620/270712

puts "Creating GitHub release for tag '#{tag}'."
puts "This is a pre-release." if prerelease

begin
  release_url = client.create_release(repo, tag, name: version, body: msg, prerelease: prerelease).rels[:self]
rescue Exception => e
  puts "A release for this tag seems to exist already (or something else went wrong)."
  puts e.message
  puts "Aborting."
  exit 2
end

asset_path = "release/XLToolbox-#{version}.exe"
puts "Uploading #{asset_path} to #{release_url.href}..."
client.upload_asset(release_url.href, asset_path)
