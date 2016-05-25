#! /usr/bin/env ruby
# Creates an XL Toolbox NG release on GitHub
# (c) 2016 Daniel Kraus (bovender)

require 'octokit'

# http://stackoverflow.com/a/11868440/270712
branch = `git rev-parse --abbrev-ref HEAD`
if branch != 'master'
  puts "Not on master. Please check out master branch first."
  exit 1
end

client = Octokit::Client.new(netrc: true)
repo = 'bovender/xltoolbox'
tag = `git describe`
msg = `git tag -n9 -l #{tag} | sed '1,2/d'`
version = tag.sub(/^v/, '')
prerelease = tag =~ /alpha|beta/

puts "Creating GitHub release for tag '#{tag}'."
puts "This is a pre-release." if prerelease

begin
  release_url = client.create_release(repo, tag, name: version, body: msg, prerelease: prerelease).rels[:self]
rescue
  puts "A release for this tag seems to exist already."
  puts "Aborting."
  exit 2
end

asset_path = "release/XLToolbox-#{version}.exe"
puts "Uploading #{asset_path} to #{release_url.href}..."
client.upload_asset(release_url.href, asset_path)
