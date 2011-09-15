#!/usr/bin/env rake
require "bundler/gem_tasks"
require 'rspec/core/rake_task'

desc "Build java jar with mvn"
task "build-jar" do
  java_excel2csv_home = "vendor/excel2csv-java"
  system "cd #{java_excel2csv_home} && mvn clean package"
  jar = FileList["#{java_excel2csv_home}/target/excel2csv*.jar"].first
  if File.exist? jar.to_s
    cp jar, "lib/excel2csv.jar"
  else
    raise "Can't find #{vendor/excel2csv-java/target}/excel2csv-x.x.x.jar. Java build is broken?"
  end
end

desc "Run specs"
RSpec::Core::RakeTask.new do |t|
  t.rspec_opts = %w(-fs --color)
end

task :default => :spec