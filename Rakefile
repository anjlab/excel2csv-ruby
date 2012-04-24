#!/usr/bin/env rake
require 'bundler/gem_tasks'
require 'rspec/core/rake_task'

desc "Run specs"
RSpec::Core::RakeTask.new do |t|
  t.rspec_opts = %w(-fs --color)
end

task :default => :spec

namespace :java do
  desc "Build java jar with mvn"
  task "build" do
    java_excel2csv_home = "vendor/excel2csv-java"
    system "cd #{java_excel2csv_home} && mvn clean package"
    jar = FileList["#{java_excel2csv_home}/target/excel2csv*.jar"].first
    if File.exist? jar.to_s
      cp jar, "lib/excel2csv.jar"
    else
      abort "Can't find #{vendor/excel2csv-java/target}/excel2csv-x.x.x.jar. Java build is broken?"
    end
  end

  desc "Pull excel2csv-java/master subtree."
  task :pull do
    if !system "git pull -s subtree java master"
      abort "Have to add java remote `git remote add -f java git@github.com:anjlab/excel2csv-java.git`"
    end
  end
end