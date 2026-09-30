#include <errno.h>
#include <signal.h>
#include <spawn.h>
#include <stdio.h>
#include <sys/wait.h>
#include <unistd.h>

extern char **environ;
static volatile sig_atomic_t child_pid = -1;

static void forward_signal(int signal_number) {
  if (child_pid > 0) kill((pid_t)child_pid, signal_number);
}

int main(int argc, char **argv) {
  if (argc < 2) {
    fputs("figma2pptx bridge launcher: missing program\n", stderr);
    return 64;
  }

  struct sigaction action = {0};
  action.sa_handler = forward_signal;
  sigemptyset(&action.sa_mask);
  sigaction(SIGTERM, &action, NULL);
  sigaction(SIGINT, &action, NULL);
  sigaction(SIGHUP, &action, NULL);

  pid_t pid;
  int spawn_error = posix_spawn(&pid, argv[1], NULL, NULL, &argv[1], environ);
  if (spawn_error != 0) {
    errno = spawn_error;
    perror("figma2pptx bridge launcher");
    return 71;
  }
  child_pid = pid;

  int status;
  while (waitpid(pid, &status, 0) < 0) {
    if (errno == EINTR) continue;
    perror("figma2pptx bridge launcher");
    return 71;
  }
  child_pid = -1;
  if (WIFEXITED(status)) return WEXITSTATUS(status);
  if (WIFSIGNALED(status)) return 128 + WTERMSIG(status);
  return 1;
}
